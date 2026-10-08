using System;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Contracts;

namespace DB2Sheet.UI
{
    /// <summary>封装支持格式化、行注释和选区执行请求的 SQL 文本编辑器。</summary>
    /// <remarks>
    /// 控件使用普通多行文本框，只接收纯文本，粘贴不会带入富文本边框。
    /// 控件不直接连接数据库。调用方负责响应 <see cref="ExecuteSelectionRequested"/> 并执行只读校验及查询。
    /// </remarks>
    public sealed class SqlEditorControl : UserControl
    {
        private readonly ISqlFormattingService _formattingService;
        private readonly TextBox _editor;
        private readonly Font _font;
        private readonly ToolStripMenuItem _copyItem;
        private readonly ToolStripMenuItem _executeSelectionItem;
        private readonly ToolStripMenuItem _formatItem;
        private readonly ToolStripMenuItem _commentItem;

        /// <summary>创建 SQL 文本编辑器。</summary>
        /// <param name="formattingService">共享的 SQL 格式化服务。</param>
        public SqlEditorControl(ISqlFormattingService formattingService)
        {
            _formattingService = formattingService ?? throw new ArgumentNullException(nameof(formattingService));
            _font = AppPresentation.CreateCodeFont();
            _editor = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                AcceptsReturn = true,
                AcceptsTab = true,
                WordWrap = false,
                ScrollBars = ScrollBars.Both,
                HideSelection = false,
                MaxLength = 0,
                Font = _font,
                BorderStyle = BorderStyle.FixedSingle
            };
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

            _editor.KeyDown += EditorKeyDown;
            _copyItem.Click += (sender, args) => _editor.Copy();
            _executeSelectionItem.Click += (sender, args) => ExecuteSelectionRequested?.Invoke(this, EventArgs.Empty);
            _formatItem.Click += (sender, args) => FormatSelectionOrAll();
            _commentItem.Click += (sender, args) => ToggleLineComments();
        }

        /// <summary>用户请求预览当前非空选区时触发。</summary>
        public event EventHandler ExecuteSelectionRequested;

        /// <summary>获取或设置编辑器中的 SQL 文本。</summary>
        public string SqlText
        {
            get => _editor.Text;
            set => _editor.Text = ToEditorNewlines(value);
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

        /// <summary>将输入焦点移入内部文本编辑器。</summary>
        public void FocusEditor()
        {
            _editor.Focus();
        }

        /// <summary>释放字体和右键菜单等托管 UI 资源。</summary>
        /// <param name="disposing">是否释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _editor.KeyDown -= EditorKeyDown;
                _editor.ContextMenuStrip?.Dispose();
                _font.Dispose();
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

        private void FormatSelectionOrAll()
        {
            bool hasSelection = _editor.SelectionLength > 0;
            int start = hasSelection ? _editor.SelectionStart : 0;
            string segment = hasSelection ? _editor.SelectedText : _editor.Text;
            string formatted = ToEditorNewlines(_formattingService.Format(segment));
            int length = hasSelection ? _editor.SelectionLength : _editor.TextLength;
            _editor.Select(start, length);
            _editor.SelectedText = formatted;
            _editor.Select(start, formatted.Length);
        }

        private void ToggleLineComments()
        {
            GetAffectedLineRange(out int start, out int end);
            string block = _editor.Text.Substring(start, end - start);
            string[] lines = SplitLinesPreservingTrailingBreak(block);
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

            ReplaceEditorText(start, end - start, string.Join("\n", lines));
        }

        private void ChangeSelectionIndent(bool outdent)
        {
            GetAffectedLineRange(out int start, out int end);
            int blockLength = Math.Max(0, end - start);
            _editor.Select(start, blockLength);
            string block = _editor.SelectedText;
            string[] lines = SplitLinesPreservingTrailingBreak(block);
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

            ReplaceEditorText(start, end - start, string.Join("\n", lines));
        }

        /// <summary>把文本写回编辑器。普通文本框只把回车换行当成换行，单独的换行符会被忽略。</summary>
        /// <param name="start">替换起点。</param>
        /// <param name="length">被替换的原文字符数。</param>
        /// <param name="value">新文本，换行可以是换行符或回车换行。</param>
        private void ReplaceEditorText(int start, int length, string value)
        {
            string replacement = ToEditorNewlines(value);
            _editor.Select(start, length);
            _editor.SelectedText = replacement;
            _editor.Select(start, SelectionLengthWithoutTrailingBreak(replacement));
        }

        /// <summary>把各种换行收成文本框要求的回车换行。</summary>
        /// <param name="value">待写入的文本。</param>
        /// <returns>只含回车换行的文本。空文本原样返回。</returns>
        private static string ToEditorNewlines(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
            return value.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "\r\n");
        }

        private static int SelectionLengthWithoutTrailingBreak(string replacement)
        {
            if (string.IsNullOrEmpty(replacement)) return 0;
            int length = replacement.Length;
            if (length >= 2 && replacement[length - 2] == '\r' && replacement[length - 1] == '\n') return length - 2;
            if (replacement[length - 1] == '\n' || replacement[length - 1] == '\r') return length - 1;
            return length;
        }

        private void GetAffectedLineRange(out int start, out int end)
        {
            int textLength = _editor.TextLength;
            if (textLength <= 0)
            {
                start = 0;
                end = 0;
                return;
            }

            int selectionStart = _editor.SelectionStart;
            int selectionLength = _editor.SelectionLength;
            int startLine = _editor.GetLineFromCharIndex(selectionStart);
            int endIndex;
            if (selectionLength <= 0)
            {
                endIndex = selectionStart;
            }
            else
            {
                int selectionEnd = selectionStart + selectionLength;
                int lastIncluded = Math.Min(selectionEnd, textLength) - 1;
                if (lastIncluded < selectionStart) lastIncluded = selectionStart;
                int endLine = _editor.GetLineFromCharIndex(lastIncluded);
                int lineStart = _editor.GetFirstCharIndexFromLine(endLine);
                if (selectionEnd == lineStart && selectionEnd > selectionStart && endLine > startLine)
                    endIndex = selectionEnd - 1;
                else
                    endIndex = lastIncluded;
            }

            int lastLine = _editor.GetLineFromCharIndex(Math.Min(Math.Max(endIndex, 0), textLength - 1));
            start = _editor.GetFirstCharIndexFromLine(startLine);
            end = lastLine + 1 < _editor.Lines.Length
                ? _editor.GetFirstCharIndexFromLine(lastLine + 1)
                : textLength;
        }

        private static string[] SplitLinesPreservingTrailingBreak(string block)
        {
            if (string.IsNullOrEmpty(block)) return new[] { string.Empty };
            return block.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        private void UpdateMenuState()
        {
            _copyItem.Enabled = _editor.SelectionLength > 0;
            _executeSelectionItem.Enabled = HasSelection;
            _formatItem.Enabled = _editor.TextLength > 0;
            _commentItem.Enabled = _editor.TextLength > 0;
        }
    }
}
