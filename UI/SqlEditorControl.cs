using System;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using WpfScrollBarVisibility = System.Windows.Controls.ScrollBarVisibility;
using System.Windows.Input;
using System.Windows.Media;
using DB2Sheet.Contracts;
using ICSharpCode.AvalonEdit;
using ICSharpCode.AvalonEdit.Document;

namespace DB2Sheet.UI
{
    /// <summary>封装支持格式化、行注释和选区执行请求的 SQL 文本编辑器。</summary>
    /// <remarks>
    /// 编辑区是 AvalonEdit，语法高亮直接使用 <see cref="SqlSyntax"/> 的 T-SQL 定义。
    /// 控件只接收纯文本。对外的 <see cref="SqlText"/> 使用回车换行，编辑器内部使用换行符。
    /// 控件不直接连接数据库。调用方负责响应 <see cref="ExecuteSelectionRequested"/> 并执行只读校验及查询。
    /// </remarks>
    public sealed class SqlEditorControl : UserControl
    {
        private readonly ISqlFormattingService _formattingService;
        private readonly ElementHost _host;
        private readonly TextEditor _editor;
        private readonly System.Drawing.Font _font;
        private readonly ContextMenuStrip _menu;
        private readonly ToolStripMenuItem _copyItem;
        private readonly ToolStripMenuItem _executeSelectionItem;
        private readonly ToolStripMenuItem _formatItem;
        private readonly ToolStripMenuItem _commentItem;

        /// <summary>创建 SQL 文本编辑器。</summary>
        /// <param name="formattingService">共享的 SQL 格式化服务。只用于排版，不负责上色。</param>
        public SqlEditorControl(ISqlFormattingService formattingService)
        {
            _formattingService = formattingService ?? throw new ArgumentNullException(nameof(formattingService));
            _font = AppPresentation.CreateCodeFont();
            _editor = new TextEditor
            {
                FontFamily = new FontFamily(_font.FontFamily.Name),
                FontSize = _font.SizeInPoints * 96.0 / 72.0,
                ShowLineNumbers = false,
                WordWrap = false,
                HorizontalScrollBarVisibility = WpfScrollBarVisibility.Auto,
                VerticalScrollBarVisibility = WpfScrollBarVisibility.Auto,
                Background = System.Windows.SystemColors.WindowBrush,
                Foreground = System.Windows.SystemColors.WindowTextBrush,
                SyntaxHighlighting = SqlSyntax.Definition
            };
            _editor.Options.ConvertTabsToSpaces = false;
            _editor.Options.CutCopyWholeLine = false;
            _editor.Options.EnableHyperlinks = false;
            _editor.Options.EnableEmailHyperlinks = false;

            _host = new ElementHost { Dock = DockStyle.Fill, Child = _editor };
            Panel border = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = System.Drawing.SystemColors.WindowFrame,
                Padding = new Padding(1)
            };
            border.Controls.Add(_host);
            Controls.Add(border);

            _copyItem = new ToolStripMenuItem("复制");
            _executeSelectionItem = new ToolStripMenuItem("执行选中文本");
            _formatItem = new ToolStripMenuItem("格式化 SQL");
            _commentItem = new ToolStripMenuItem("注释/取消注释");
            _menu = new ContextMenuStrip();
            _menu.Items.Add(_copyItem);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(_executeSelectionItem);
            _menu.Items.Add(_formatItem);
            _menu.Items.Add(_commentItem);
            _menu.Opening += (sender, args) => UpdateMenuState();

            _editor.TextChanged += EditorTextChanged;
            _editor.TextArea.PreviewKeyDown += EditorPreviewKeyDown;
            _editor.TextArea.PreviewMouseRightButtonDown += EditorPreviewMouseRightButtonDown;
            _copyItem.Click += (sender, args) => _editor.Copy();
            _executeSelectionItem.Click += (sender, args) => ExecuteSelectionRequested?.Invoke(this, EventArgs.Empty);
            _formatItem.Click += (sender, args) => FormatSelectionOrAll();
            _commentItem.Click += (sender, args) => ToggleLineComments();
        }

        /// <summary>用户请求预览当前非空选区时触发。</summary>
        public event EventHandler ExecuteSelectionRequested;

        /// <summary>编辑器文本变化时触发，供页签刷新未保存标记。</summary>
        /// <remarks>设置 <see cref="SqlText"/> 也会触发。调用方在写入基线文本后再订阅，避免把载入当成修改。</remarks>
        public event EventHandler SqlTextChanged;

        /// <summary>获取或设置编辑器中的 SQL 文本。</summary>
        /// <remarks>读取时把内部换行符收成回车换行。写入时由编辑器自己规范化换行。</remarks>
        public string SqlText
        {
            get => ToExternalNewlines(_editor.Text);
            set => _editor.Text = value ?? string.Empty;
        }

        /// <summary>获取当前选中的 SQL 文本。</summary>
        public string SelectedSql => ToExternalNewlines(_editor.SelectedText);

        /// <summary>获取当前是否存在非空文本选区。</summary>
        public bool HasSelection => _editor.SelectionLength > 0 && !string.IsNullOrWhiteSpace(_editor.SelectedText);

        /// <summary>清空 SQL 并将输入焦点移入编辑器。</summary>
        public void ClearAndFocus()
        {
            _editor.Text = string.Empty;
            FocusEditor();
        }

        /// <summary>将输入焦点移入内部文本编辑器。</summary>
        public void FocusEditor()
        {
            _host.Focus();
            _editor.Focus();
            _editor.TextArea.Focus();
        }

        /// <summary>释放字体、宿主和右键菜单等托管 UI 资源。</summary>
        /// <param name="disposing">是否释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _editor.TextChanged -= EditorTextChanged;
                _editor.TextArea.PreviewKeyDown -= EditorPreviewKeyDown;
                _editor.TextArea.PreviewMouseRightButtonDown -= EditorPreviewMouseRightButtonDown;
                _menu.Dispose();
                _host.Dispose();
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

        private void EditorTextChanged(object sender, EventArgs e)
        {
            SqlTextChanged?.Invoke(this, EventArgs.Empty);
        }

        private void EditorPreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            System.Windows.Input.ModifierKeys modifiers = Keyboard.Modifiers;
            bool control = (modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control;
            bool shift = (modifiers & System.Windows.Input.ModifierKeys.Shift) == System.Windows.Input.ModifierKeys.Shift;
            if (control && !shift && (e.Key == Key.Oem2 || e.Key == Key.Divide))
            {
                ToggleLineComments();
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Tab || control) return;

            ChangeSelectionIndent(shift);
            e.Handled = true;
        }

        /// <summary>右键弹出原有菜单，不移动插入点，以便选区执行仍针对右键前的选区。</summary>
        /// <param name="sender">文本区。</param>
        /// <param name="e">鼠标事件。标记为已处理，避免编辑器改选区。</param>
        private void EditorPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            FocusEditor();
            _menu.Show(System.Windows.Forms.Cursor.Position);
        }

        private void FormatSelectionOrAll()
        {
            bool hasSelection = _editor.SelectionLength > 0;
            int start = hasSelection ? _editor.SelectionStart : 0;
            string segment = hasSelection ? _editor.SelectedText : _editor.Text;
            string formatted = ToDocumentNewlines(_formattingService.Format(segment));
            int length = hasSelection ? _editor.SelectionLength : _editor.Document.TextLength;
            _editor.Document.Replace(start, length, formatted);
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

            ReplaceDocument(start, end - start, string.Join("\n", lines));
        }

        private void ChangeSelectionIndent(bool outdent)
        {
            GetAffectedLineRange(out int start, out int end);
            int blockLength = Math.Max(0, end - start);
            string block = _editor.Document.GetText(start, blockLength);
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

            ReplaceDocument(start, end - start, string.Join("\n", lines));
        }

        /// <summary>替换文档中的一段，并选中替换结果里除末尾换行以外的部分。</summary>
        /// <param name="start">替换起点，按编辑器内部偏移。</param>
        /// <param name="length">被替换的原文字符数。</param>
        /// <param name="value">新文本。换行会收成换行符。</param>
        private void ReplaceDocument(int start, int length, string value)
        {
            string replacement = ToDocumentNewlines(value);
            _editor.Document.Replace(start, length, replacement);
            _editor.Select(start, SelectionLengthWithoutTrailingBreak(replacement));
        }

        /// <summary>把各种换行收成编辑器内部使用的换行符。</summary>
        /// <param name="value">待写入的文本。</param>
        /// <returns>只含换行符的文本。空文本原样返回。</returns>
        private static string ToDocumentNewlines(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
            return value.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        /// <summary>把编辑器内部的换行符收成对外的回车换行。</summary>
        /// <param name="value">编辑器文本。</param>
        /// <returns>只含回车换行的文本。空文本原样返回。</returns>
        private static string ToExternalNewlines(string value)
        {
            if (string.IsNullOrEmpty(value)) return value ?? string.Empty;
            return ToDocumentNewlines(value).Replace("\n", "\r\n");
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
            TextDocument document = _editor.Document;
            int textLength = document.TextLength;
            if (textLength <= 0)
            {
                start = 0;
                end = 0;
                return;
            }

            int selectionStart = _editor.SelectionStart;
            int selectionLength = _editor.SelectionLength;
            DocumentLine startLine = document.GetLineByOffset(selectionStart);
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
                DocumentLine endLine = document.GetLineByOffset(lastIncluded);
                if (selectionEnd == endLine.Offset && selectionEnd > selectionStart && endLine.LineNumber > startLine.LineNumber)
                    endIndex = selectionEnd - 1;
                else
                    endIndex = lastIncluded;
            }

            DocumentLine lastLine = document.GetLineByOffset(Math.Min(Math.Max(endIndex, 0), textLength - 1));
            start = startLine.Offset;
            end = lastLine.Offset + lastLine.TotalLength;
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
            _formatItem.Enabled = _editor.Document.TextLength > 0;
            _commentItem.Enabled = _editor.Document.TextLength > 0;
        }
    }
}
