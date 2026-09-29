using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.UI
{
    /// <summary>封装支持语法高亮、格式化、行注释和选区执行请求的 SQL 富文本编辑器。</summary>
    /// <remarks>
    /// 控件只处理文本和 UI 状态，不直接连接数据库。高亮在 UI 线程由短延迟计时器执行；
    /// 调用方负责响应 <see cref="ExecuteSelectionRequested"/> 并执行只读校验及查询。
    /// </remarks>
    public sealed class SqlEditorControl : UserControl
    {
        private const int EmGetScrollPos = 0x04DD;
        private const int EmSetScrollPos = 0x04DE;
        private const int WmSetRedraw = 0x000B;

        private readonly ISqlFormattingService _formattingService;
        private readonly RichTextBox _editor;
        private readonly Timer _highlightTimer;
        private readonly Font _regularFont;
        private readonly Font _keywordFont;
        private readonly ToolStripMenuItem _copyItem;
        private readonly ToolStripMenuItem _executeSelectionItem;
        private readonly ToolStripMenuItem _formatItem;
        private readonly ToolStripMenuItem _commentItem;
        private bool _applyingHighlight;

        /// <summary>创建 SQL 富文本编辑器。</summary>
        /// <param name="formattingService">共享的 SQL 词法分析和格式化服务。</param>
        public SqlEditorControl(ISqlFormattingService formattingService)
        {
            _formattingService = formattingService ?? throw new ArgumentNullException(nameof(formattingService));
            _regularFont = new Font(AppPresentation.CodeFontName, AppPresentation.CodeFontSize, FontStyle.Regular);
            _keywordFont = new Font(AppPresentation.CodeFontName, AppPresentation.CodeFontSize, FontStyle.Bold);

            _editor = new RichTextBox
            {
                Dock = DockStyle.Fill,
                AcceptsTab = true,
                DetectUrls = false,
                Font = _regularFont,
                WordWrap = false,
                BorderStyle = BorderStyle.FixedSingle,
                HideSelection = false
            };
            _highlightTimer = new Timer { Interval = 250 };
            _copyItem = new ToolStripMenuItem("复制");
            _executeSelectionItem = new ToolStripMenuItem("执行选中文本");
            _formatItem = new ToolStripMenuItem("格式化 SQL");
            _commentItem = new ToolStripMenuItem("注释/取消注释");

            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add(_copyItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(_executeSelectionItem);
            menu.Items.Add(_formatItem);
            menu.Items.Add(_commentItem);
            menu.Opening += (sender, args) => UpdateMenuState();
            _editor.ContextMenuStrip = menu;
            Controls.Add(_editor);

            _editor.TextChanged += EditorTextChanged;
            _editor.KeyDown += EditorKeyDown;
            _highlightTimer.Tick += HighlightTimerTick;
            _copyItem.Click += (sender, args) => _editor.Copy();
            _executeSelectionItem.Click += (sender, args) => ExecuteSelectionRequested?.Invoke(this, EventArgs.Empty);
            _formatItem.Click += (sender, args) => FormatSelectionOrAll();
            _commentItem.Click += (sender, args) => ToggleLineComments();
        }

        /// <summary>用户请求预览当前非空选区时触发。</summary>
        public event EventHandler ExecuteSelectionRequested;

        /// <summary>获取或设置编辑器中的纯 SQL 文本。</summary>
        public string SqlText
        {
            get => _editor.Text;
            set
            {
                _editor.Text = value ?? string.Empty;
                ApplyHighlight();
            }
        }

        /// <summary>获取当前选中的 SQL 文本。</summary>
        public string SelectedSql => _editor.SelectedText;

        /// <summary>获取当前是否存在非空文本选区。</summary>
        public bool HasSelection => _editor.SelectionLength > 0 && !string.IsNullOrWhiteSpace(_editor.SelectedText);

        /// <summary>清空 SQL 并将输入焦点移入编辑器。</summary>
        public void ClearAndFocus()
        {
            _editor.Clear();
            _editor.Focus();
        }

        /// <summary>将输入焦点移入内部富文本编辑器。</summary>
        public void FocusEditor()
        {
            _editor.Focus();
        }

        /// <summary>释放计时器、字体和右键菜单等托管 UI 资源。</summary>
        /// <param name="disposing">是否释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _highlightTimer.Stop();
                _highlightTimer.Dispose();
                _editor.KeyDown -= EditorKeyDown;
                _editor.ContextMenuStrip?.Dispose();
                _keywordFont.Dispose();
                _regularFont.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>格式化当前选区；若无选区则格式化全部 SQL。</summary>
        public void FormatSql()
        {
            FormatSelectionOrAll();
        }

        /// <summary>对当前行或所选行执行注释/取消注释切换。</summary>
        public void ToggleComment()
        {
            ToggleLineComments();
        }

        private void EditorTextChanged(object sender, EventArgs e)
        {
            if (_applyingHighlight) return;
            _highlightTimer.Stop();
            _highlightTimer.Start();
        }

        private void HighlightTimerTick(object sender, EventArgs e)
        {
            _highlightTimer.Stop();
            ApplyHighlight();
        }

        private void EditorKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && (e.KeyCode == Keys.OemQuestion || e.KeyCode == Keys.Oem2))
            {
                ToggleLineComments();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.KeyCode != Keys.Tab) return;

            ChangeSelectionIndent(e.Shift);
            e.Handled = true;
            e.SuppressKeyPress = true;
        }

        private void ApplyHighlight()
        {
            if (_editor.IsDisposed || !_editor.IsHandleCreated) return;
            int selectionStart = _editor.SelectionStart;
            int selectionLength = _editor.SelectionLength;
            Point scrollPosition = new Point();
            SendMessage(_editor.Handle, EmGetScrollPos, IntPtr.Zero, ref scrollPosition);
            _applyingHighlight = true;
            SendMessage(_editor.Handle, WmSetRedraw, IntPtr.Zero, IntPtr.Zero);
            try
            {
                _editor.SelectAll();
                _editor.SelectionColor = SystemColors.WindowText;
                _editor.SelectionFont = _regularFont;
                foreach (SqlToken token in _formattingService.Tokenize(_editor.Text))
                {
                    _editor.Select(token.Start, token.Length);
                    _editor.SelectionColor = ColorFor(token.Kind);
                    _editor.SelectionFont = token.Kind == SqlTokenKind.Keyword ? _keywordFont : _regularFont;
                }
                _editor.Select(Math.Min(selectionStart, _editor.TextLength), Math.Min(selectionLength, Math.Max(0, _editor.TextLength - selectionStart)));
                SendMessage(_editor.Handle, EmSetScrollPos, IntPtr.Zero, ref scrollPosition);
            }
            finally
            {
                SendMessage(_editor.Handle, WmSetRedraw, new IntPtr(1), IntPtr.Zero);
                _editor.Invalidate();
                _applyingHighlight = false;
            }
        }

        private void FormatSelectionOrAll()
        {
            int start = _editor.SelectionLength > 0 ? _editor.SelectionStart : 0;
            int length = _editor.SelectionLength > 0 ? _editor.SelectionLength : _editor.TextLength;
            string formatted = _formattingService.Format(_editor.Text.Substring(start, length));
            _editor.Select(start, length);
            _editor.SelectedText = formatted;
            _editor.Select(start, formatted.Length);
            ApplyHighlight();
        }

        private void ToggleLineComments()
        {
            int originalStart = _editor.SelectionStart;
            int originalEnd = originalStart + Math.Max(1, _editor.SelectionLength);
            int start = _editor.GetFirstCharIndexFromLine(_editor.GetLineFromCharIndex(originalStart));
            int endLine = _editor.GetLineFromCharIndex(Math.Min(originalEnd, Math.Max(0, _editor.TextLength - 1)));
            int end = endLine + 1 < _editor.Lines.Length
                ? _editor.GetFirstCharIndexFromLine(endLine + 1)
                : _editor.TextLength;
            string block = _editor.Text.Substring(start, end - start);
            string[] lines = block.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            bool uncomment = true;
            foreach (string line in lines)
            {
                if (line.Length > 0 && !line.TrimStart().StartsWith("--", StringComparison.Ordinal))
                {
                    uncomment = false;
                    break;
                }
            }

            for (int index = 0; index < lines.Length; index++)
            {
                if (lines[index].Length == 0) continue;
                int content = lines[index].Length - lines[index].TrimStart().Length;
                if (uncomment)
                {
                    if (lines[index].Substring(content).StartsWith("-- ", StringComparison.Ordinal))
                        lines[index] = lines[index].Remove(content, 3);
                    else if (lines[index].Substring(content).StartsWith("--", StringComparison.Ordinal))
                        lines[index] = lines[index].Remove(content, 2);
                }
                else
                {
                    lines[index] = lines[index].Insert(content, "-- ");
                }
            }

            string replacement = string.Join(Environment.NewLine, lines);
            _editor.Select(start, end - start);
            _editor.SelectedText = replacement;
            _editor.Select(start, replacement.Length);
            ApplyHighlight();
        }

        private void ChangeSelectionIndent(bool outdent)
        {
            int originalStart = _editor.SelectionStart;
            int originalEnd = originalStart + Math.Max(1, _editor.SelectionLength);
            int start = _editor.GetFirstCharIndexFromLine(_editor.GetLineFromCharIndex(originalStart));
            int endLine = _editor.GetLineFromCharIndex(Math.Min(originalEnd, Math.Max(0, _editor.TextLength - 1)));
            int end = endLine + 1 < _editor.Lines.Length
                ? _editor.GetFirstCharIndexFromLine(endLine + 1)
                : _editor.TextLength;

            string block = _editor.Text.Substring(start, end - start);
            string[] lines = block.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int index = 0; index < lines.Length; index++)
            {
                if (lines[index].Length == 0) continue;

                if (outdent)
                {
                    if (lines[index].StartsWith("\t", StringComparison.Ordinal)) lines[index] = lines[index].Substring(1);
                    else if (lines[index].StartsWith("    ", StringComparison.Ordinal)) lines[index] = lines[index].Substring(4);
                }
                else
                {
                    lines[index] = "\t" + lines[index];
                }
            }

            string replacement = string.Join(Environment.NewLine, lines);
            _editor.Select(start, end - start);
            _editor.SelectedText = replacement;
            _editor.Select(start, replacement.Length);
            ApplyHighlight();
        }

        private void UpdateMenuState()
        {
            _copyItem.Enabled = _editor.SelectionLength > 0;
            _executeSelectionItem.Enabled = HasSelection;
            _formatItem.Enabled = _editor.TextLength > 0;
            _commentItem.Enabled = _editor.TextLength > 0;
        }

        private static Color ColorFor(SqlTokenKind kind)
        {
            switch (kind)
            {
                case SqlTokenKind.Keyword: return Color.RoyalBlue;
                case SqlTokenKind.String: return Color.Firebrick;
                case SqlTokenKind.Comment: return Color.ForestGreen;
                case SqlTokenKind.Number: return Color.DarkCyan;
                case SqlTokenKind.Parameter: return Color.DarkMagenta;
                case SqlTokenKind.QuotedIdentifier: return Color.SaddleBrown;
                default: return SystemColors.WindowText;
            }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr handle, int message, IntPtr wParam, ref Point lParam);
    }
}
