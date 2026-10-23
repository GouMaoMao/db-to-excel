using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using DB2Sheet.Infrastructure;
using DB2Sheet.Providers;

namespace DB2Sheet.UI
{
    /// <summary>列出连接、查询、Skill 等本机文件的实际路径。</summary>
    /// <remarks>
    /// 只展示 <see cref="ApplicationPaths"/> 算出的位置，不搬文件，也不改保存目录。
    /// 小窗不写入 <c>session.formBounds</c>。必须在界面线程以模态方式打开。
    /// </remarks>
    public sealed class LocalDataForm : AppForm
    {
        private readonly ToolTip _tips = new ToolTip();
        private readonly ListView _list;
        private readonly Button _copyButton;
        private readonly Button _openButton;
        private bool _fittingColumn;
        private ListViewItem _tipItem;

        /// <summary>创建本地数据窗，并按固定顺序填入路径。</summary>
        /// <param name="paths">当前用户的路径集合。为空时无法知道文件在哪。</param>
        public LocalDataForm(ApplicationPaths paths)
        {
            if (paths == null) throw new ArgumentNullException(nameof(paths));

            Text = AppPresentation.WindowTitle("本地数据");
            Width = 720;
            Height = 440;
            MinimumSize = new Size(520, 320);
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            MinimizeBox = false;
            CenterOnWorkArea();

            Label intro = new Label
            {
                Text = AppPresentation.LocalDataIntro,
                AutoSize = true,
                Padding = new Padding(0, 0, 0, 8)
            };

            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            _list.Columns.Add("项目", 140);
            _list.Columns.Add("路径", 520);
            AddItem("数据目录", paths.RootDirectory, true, AppPresentation.LocalDataRootTip);
            AddItem("连接方案", paths.ConnectionsFile, false, AppPresentation.LocalDataConnectionsTip);
            AddItem("查询方案", paths.QueryProfilesDirectory, true, AppPresentation.LocalDataQueriesTip);
            AddItem("参数设置", paths.SettingsFile, false, AppPresentation.LocalDataSettingsTip);
            AddItem("JDBC 环境", paths.JdbcEnvironmentFile, false, AppPresentation.LocalDataJdbcTip);
            AddItem("模型配置", paths.AiModelsFile, false, AppPresentation.LocalDataModelsTip);
            AddItem("AI 思考强度映射", paths.AiThinkingMappingsFile, false, AppPresentation.LocalDataThinkingMappingsTip);
            AddItem("Skill", paths.AiSkillsDirectory, true, AppPresentation.LocalDataSkillsTip);
            AddItem("对话", paths.AiConversationsDirectory, true, AppPresentation.LocalDataConversationsTip);
            AddItem("日志", paths.LogsDirectory, true, AppPresentation.LocalDataLogsTip);
            AddItem("对话页缓存", paths.WebView2UserDataDirectory, true, AppPresentation.LocalDataWebViewTip);
            AddItem(
                "JDBC 转接程序",
                Path.Combine(paths.RootDirectory, JdbcBridgeFiles.FileName),
                false,
                AppPresentation.LocalDataBridgeTip);
            if (_list.Items.Count > 0)
                _list.Items[0].Selected = true;

            _copyButton = new Button { Text = "复制路径", AutoSize = true, UseVisualStyleBackColor = true };
            _openButton = new Button { Text = "打开位置", AutoSize = true, UseVisualStyleBackColor = true };
            Button closeButton = new Button
            {
                Text = "关闭",
                AutoSize = true,
                UseVisualStyleBackColor = true,
                DialogResult = DialogResult.Cancel
            };
            _tips.SetToolTip(_copyButton, AppPresentation.LocalDataCopyTip);
            _tips.SetToolTip(_openButton, AppPresentation.LocalDataOpenTip);
            _tips.SetToolTip(closeButton, AppPresentation.LocalDataCloseTip);

            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Padding = new Padding(0, 8, 0, 0)
            };
            footer.Controls.Add(closeButton);
            footer.Controls.Add(_openButton);
            footer.Controls.Add(_copyButton);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 3
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(intro, 0, 0);
            root.Controls.Add(_list, 0, 1);
            root.Controls.Add(footer, 0, 2);
            Controls.Add(root);

            _list.SelectedIndexChanged += (sender, args) => UpdateActions();
            _list.Resize += (sender, args) => FitPathColumn();
            _list.MouseMove += ListMouseMove;
            _list.MouseLeave += (sender, args) => ShowRowTip(null);
            _copyButton.Click += CopySelected;
            _openButton.Click += OpenSelected;
            CancelButton = closeButton;
            Load += (sender, args) => FitPathColumn();
            UpdateActions();
        }

        /// <summary>释放悬浮提示。</summary>
        /// <param name="disposing">为 true 时释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _tips.Dispose();
            base.Dispose(disposing);
        }

        /// <summary>加入一行。路径来自 <see cref="ApplicationPaths"/>，不在这里拼文件名，转接程序除外。</summary>
        /// <param name="name">列表里的项目名。</param>
        /// <param name="path">完整路径。文件尚未生成时仍然显示。</param>
        /// <param name="isDirectory">为 true 时按目录打开，否则在资源管理器中定位文件。</param>
        /// <param name="tip">悬浮说明。第一行是概述。</param>
        private void AddItem(string name, string path, bool isDirectory, string tip)
        {
            ListViewItem item = new ListViewItem(name);
            item.SubItems.Add(path ?? string.Empty);
            item.ToolTipText = tip ?? string.Empty;
            item.Tag = new LocalDataItem(name, path ?? string.Empty, isDirectory);
            _list.Items.Add(item);
        }

        /// <summary>鼠标停在某一行时显示该项说明。路径列和名称列用同一段文字。</summary>
        /// <param name="sender">列表。</param>
        /// <param name="e">相对列表客户区的坐标。</param>
        private void ListMouseMove(object sender, MouseEventArgs e)
        {
            ShowRowTip(_list.HitTest(e.X, e.Y).Item);
        }

        /// <summary>换行时才重设悬浮说明，避免每次移动都重置出现时间。</summary>
        /// <param name="item">当前行。空白处为空。</param>
        private void ShowRowTip(ListViewItem item)
        {
            if (ReferenceEquals(item, _tipItem)) return;
            _tipItem = item;
            _tips.SetToolTip(_list, item == null ? string.Empty : (item.ToolTipText ?? string.Empty));
        }

        /// <summary>没有选中项时不能复制或打开。</summary>
        private void UpdateActions()
        {
            bool selected = SelectedItem() != null;
            _copyButton.Enabled = selected;
            _openButton.Enabled = selected;
        }

        /// <summary>把选中路径放入剪贴板。</summary>
        /// <remarks>剪贴板被其他程序占用时复制失败，窗口仍留着供查看。</remarks>
        private void CopySelected(object sender, EventArgs e)
        {
            LocalDataItem item = SelectedItem();
            if (item == null || item.Path.Length == 0) return;
            try
            {
                Clipboard.SetText(item.Path);
            }
            catch (Exception)
            {
                // 剪贴板偶发占用；复制失败不阻断继续查看路径。
            }
        }

        /// <summary>在资源管理器中打开选中项。尚未生成时只说明，不创建文件或目录。</summary>
        /// <remarks>在界面线程启动资源管理器。启动失败时弹出错误详情，不改任何本地文件。</remarks>
        private void OpenSelected(object sender, EventArgs e)
        {
            LocalDataItem item = SelectedItem();
            if (item == null || item.Path.Length == 0) return;
            bool exists = item.IsDirectory ? Directory.Exists(item.Path) : File.Exists(item.Path);
            if (!exists)
            {
                MessageBox.Show(
                    this,
                    "「" + item.Name + "」还没有生成。路径仍可复制。",
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = item.IsDirectory
                        ? "\"" + item.Path + "\""
                        : "/select,\"" + item.Path + "\"",
                    UseShellExecute = true
                });
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
            }
        }

        /// <summary>返回当前选中项。没有选中时为空。</summary>
        private LocalDataItem SelectedItem()
        {
            if (_list.SelectedItems.Count == 0) return null;
            return _list.SelectedItems[0].Tag as LocalDataItem;
        }

        /// <summary>路径列至少铺满列表；比窗口更长时保留横向滚动，避免路径被截断后看不到。</summary>
        /// <remarks>改列宽会同步触发列表的尺寸变化。用标记挡住重入，避免滚动条出现和消失时来回改宽度。</remarks>
        private void FitPathColumn()
        {
            if (_fittingColumn || _list.Columns.Count < 2 || _list.ClientSize.Width <= 0) return;
            int width = _list.ClientSize.Width - _list.Columns[0].Width - 8;
            if (width < 240) width = 240;
            foreach (ListViewItem item in _list.Items)
            {
                if (item.SubItems.Count < 2) continue;
                int textWidth = TextRenderer.MeasureText(item.SubItems[1].Text, _list.Font).Width + 16;
                if (textWidth > width) width = textWidth;
            }

            if (_list.Columns[1].Width == width) return;
            _fittingColumn = true;
            try
            {
                _list.Columns[1].Width = width;
            }
            finally
            {
                _fittingColumn = false;
            }
        }

        /// <summary>相对光标所在屏幕的工作区居中。不写入会话设置。</summary>
        /// <remarks>在创建句柄之前设置位置，避免先画出角落再跳到中间。</remarks>
        private void CenterOnWorkArea()
        {
            Rectangle area = FormSizeMemory.CursorWorkingArea();
            Location = new Point(
                area.Left + Math.Max(0, (area.Width - Width) / 2),
                area.Top + Math.Max(0, (area.Height - Height) / 2));
        }

        /// <summary>列表中的一项：名称、路径，以及应按文件还是目录打开。</summary>
        private sealed class LocalDataItem
        {
            /// <summary>记下这一行要展示和打开的路径。</summary>
            /// <param name="name">用户可见名称。</param>
            /// <param name="path">完整路径。</param>
            /// <param name="isDirectory">为 true 时打开目录，否则定位文件。</param>
            public LocalDataItem(string name, string path, bool isDirectory)
            {
                Name = name ?? string.Empty;
                Path = path ?? string.Empty;
                IsDirectory = isDirectory;
            }

            /// <summary>列表中的项目名。</summary>
            public string Name { get; }

            /// <summary>完整路径。文件可以尚未生成。</summary>
            public string Path { get; }

            /// <summary>为 true 时这是目录。</summary>
            public bool IsDirectory { get; }
        }
    }
}
