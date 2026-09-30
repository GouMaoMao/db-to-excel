using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Providers;

namespace DB2Sheet.UI
{
    /// <summary>让用户检测或指定本机 java.exe，供 JDBC 查询使用。</summary>
    /// <remarks>
    /// 厂商驱动不在此配置，而在连接方案中选择。窗体记住宽高。
    /// 「检测 Java」会搜索本机常见位置并填入路径；不提供一键下载安装。
    /// </remarks>
    public sealed class JdbcEnvironmentForm : AppForm
    {
        private readonly IJdbcEnvironmentStore _store;
        private readonly TextBox _javaPath;
        private readonly Button _browseJava;
        private readonly Button _detectJavaButton;
        private readonly Button _openDownloadButton;
        private readonly Label _javaStatus;
        private readonly List<Label> _guideLabels = new List<Label>();

        /// <summary>创建 JDBC 环境窗体。</summary>
        /// <param name="store">环境设置存储。</param>
        /// <param name="settings">用于记住窗体尺寸的设置存储。</param>
        public JdbcEnvironmentForm(IJdbcEnvironmentStore store, ISettingsStore settings)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            Text = AppPresentation.WindowTitle("JDBC 环境");
            Width = 720;
            Height = 520;
            MinimumSize = new Size(560, 460);
            StartPosition = FormStartPosition.Manual;
            FormSizeMemory.Attach(this, settings, "JdbcEnvironment");

            JdbcEnvironmentSettings current = _store.Load();
            _javaPath = new TextBox { Dock = DockStyle.Fill, Text = current.JavaExecutable ?? string.Empty };
            _browseJava = new Button { Text = "浏览…", AutoSize = true };
            _detectJavaButton = new Button { Text = "检测 Java", AutoSize = true };
            _openDownloadButton = new Button { Text = "打开 JDK 下载页", AutoSize = true };
            _javaStatus = CreateStatusLabel();

            GroupBox javaGroup = new GroupBox
            {
                Text = "Java",
                Dock = DockStyle.Fill,
                Padding = new Padding(8)
            };
            TableLayoutPanel javaLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                AutoSize = true
            };
            javaLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int index = 0; index < 6; index++)
            {
                javaLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            Control guide = CreateGuide();
            guide.SizeChanged += GuideSizeChanged;
            javaLayout.Controls.Add(guide, 0, 0);
            FlowLayoutPanel detectActions = CreateActions();
            detectActions.Controls.Add(_detectJavaButton);
            javaLayout.Controls.Add(detectActions, 0, 1);
            javaLayout.Controls.Add(CreatePathRow(_javaPath, _browseJava), 0, 2);
            FlowLayoutPanel downloadActions = CreateActions();
            downloadActions.Controls.Add(_openDownloadButton);
            javaLayout.Controls.Add(downloadActions, 0, 3);
            javaLayout.Controls.Add(_javaStatus, 0, 4);
            javaGroup.Controls.Add(javaLayout);

            Button saveButton = new Button { Text = "保存", AutoSize = true };
            Button closeButton = new Button { Text = "关闭", AutoSize = true, DialogResult = DialogResult.Cancel };
            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true
            };
            footer.Controls.Add(closeButton);
            footer.Controls.Add(saveButton);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 2
            };
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(javaGroup, 0, 0);
            root.Controls.Add(footer, 0, 1);
            Controls.Add(root);

            _browseJava.Click += BrowseJava;
            _detectJavaButton.Click += DetectJava;
            _openDownloadButton.Click += (sender, args) => Process.Start(JavaRuntimeProbe.DownloadUrl);
            saveButton.Click += SaveClick;
            AcceptButton = saveButton;
            CancelButton = closeButton;
        }

        private async void DetectJava(object sender, EventArgs e)
        {
            try
            {
                SetBusy(true);
                _javaStatus.Text = "正在搜索本机 Java…";
                string preferred = _javaPath.Text.Trim();
                JavaResolution resolution = await Task.Run(() => JavaRuntimeProbe.Discover(preferred)).ConfigureAwait(true);
                if (!resolution.Found)
                {
                    SetStatus(_javaStatus, false, resolution.Detail);
                    return;
                }

                _javaPath.Text = resolution.Executable;
                SaveCore();
                string version = await Task.Run(() => JavaRuntimeProbe.TryReadVersion(resolution.Executable)).ConfigureAwait(true);
                string message = string.IsNullOrWhiteSpace(version)
                    ? "已找到 " + resolution.Executable + "，但无法读取版本。请确认它是可用的 java.exe。"
                    : "已找到 " + version + Environment.NewLine + resolution.Executable;
                SetStatus(_javaStatus, !string.IsNullOrWhiteSpace(version), message);
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
                SetStatus(_javaStatus, false, exception.Message);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void SaveClick(object sender, EventArgs e)
        {
            try
            {
                SaveCore();
                JavaResolution resolution = JavaRuntimeProbe.Resolve(_javaPath.Text.Trim());
                if (resolution.Found &&
                    !string.Equals(_javaPath.Text.Trim(), resolution.Executable, StringComparison.OrdinalIgnoreCase))
                {
                    _javaPath.Text = resolution.Executable;
                    SaveCore();
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(this, exception);
                SetStatus(_javaStatus, false, exception.Message);
            }
        }

        private void SaveCore()
        {
            _store.Save(new JdbcEnvironmentSettings
            {
                JavaExecutable = _javaPath.Text.Trim()
            });
        }

        private void BrowseJava(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "java.exe|java.exe|所有文件|*.*",
                FileName = _javaPath.Text
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _javaPath.Text = dialog.FileName;
                }
            }
        }

        private void SetBusy(bool busy)
        {
            _detectJavaButton.Enabled = !busy;
            _browseJava.Enabled = !busy;
            _javaPath.Enabled = !busy;
            _openDownloadButton.Enabled = !busy;
            UseWaitCursor = busy;
        }

        private static void SetStatus(Label label, bool succeeded, string message)
        {
            label.ForeColor = succeeded ? Color.DarkGreen : Color.DarkRed;
            label.Text = message ?? string.Empty;
        }

        private static Label CreateStatusLabel()
        {
            return new Label
            {
                AutoSize = true,
                Dock = DockStyle.Top,
                MaximumSize = new Size(1000, 0),
                MinimumSize = new Size(0, 40),
                ForeColor = SystemColors.GrayText,
                Padding = new Padding(0, 8, 0, 0)
            };
        }

        /// <summary>拼出三段引导：为何需要 Java、何时安装、以及逐步操作。步骤各占一行。</summary>
        /// <returns>随窗体宽度换行的说明面板。调用方把它放进 Java 分组，并在尺寸变化时收窄标签。</returns>
        private Control CreateGuide()
        {
            TableLayoutPanel guide = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 6,
                Margin = new Padding(0, 0, 0, 4)
            };
            guide.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int index = 0; index < 6; index++)
            {
                guide.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            AddGuideSection(guide, 0, "为什么需要 Java？",
                "只有 JDBC 连接需要本机 Java（JDK 8 或更高）。使用 MySQL、SQL Server、PostgreSQL、SQLite 等内置连接时，不必安装 Java。",
                first: true);
            AddGuideSection(guide, 2, "什么时候需要安装？",
                "连接方式是 JDBC，且本机还没有 Java；或者已有 Java，但版本低于 JDK 8。",
                first: false);
            AddGuideSection(guide, 4, "操作步骤：",
                "1. 点击「检测 Java」，自动搜索本机已安装的 Java。" + Environment.NewLine +
                "2. 若未找到，可点「浏览…」指定 java.exe，或点「打开 JDK 下载页」自行安装。" + Environment.NewLine +
                "3. 安装完成后，再点一次「检测 Java」。",
                first: false);
            return guide;
        }

        /// <summary>向引导面板追加一个加粗小标题和一段正文。小标题使用独立字体，由标签在释放时一并释放。</summary>
        /// <param name="guide">引导面板。</param>
        /// <param name="row">小标题所在行；正文占用下一行。</param>
        /// <param name="title">加粗小标题。</param>
        /// <param name="body">正文。操作步骤在文本内用换行分成多行。</param>
        /// <param name="first">是否为第一段。后续段落与上一段之间留出空行。</param>
        private void AddGuideSection(TableLayoutPanel guide, int row, string title, string body, bool first)
        {
            Label heading = CreateWrappingLabel(title);
            heading.Font = new Font(Font, FontStyle.Bold);
            heading.Margin = new Padding(0, first ? 0 : 10, 0, 2);
            Label text = CreateWrappingLabel(body);
            text.Margin = new Padding(0, 0, 0, 2);
            guide.Controls.Add(heading, 0, row);
            guide.Controls.Add(text, 0, row + 1);
        }

        /// <summary>创建会随引导面板宽度折行的标签，并登记以便在面板变窄时更新最大宽度。</summary>
        /// <param name="text">标签文本。</param>
        /// <returns>自动高度、限制宽度的标签。</returns>
        private Label CreateWrappingLabel(string text)
        {
            Label label = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(640, 0),
                Text = text
            };
            _guideLabels.Add(label);
            return label;
        }

        /// <summary>把引导文字的最大宽度收成面板客户区宽度，避免长句横向撑出窗体。</summary>
        private void GuideSizeChanged(object sender, EventArgs e)
        {
            Control host = sender as Control;
            if (host == null || host.ClientSize.Width <= 0) return;
            int width = host.ClientSize.Width;
            foreach (Label label in _guideLabels)
            {
                if (label.MaximumSize.Width == width) continue;
                label.MaximumSize = new Size(width, 0);
            }
        }

        private static Control CreatePathRow(Control editor, params Button[] buttons)
        {
            TableLayoutPanel row = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 1 + (buttons == null ? 0 : buttons.Length),
                RowCount = 1
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.Controls.Add(editor, 0, 0);
            int column = 1;
            if (buttons != null)
            {
                foreach (Button button in buttons)
                {
                    row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                    row.Controls.Add(button, column, 0);
                    column++;
                }
            }

            return row;
        }

        private static FlowLayoutPanel CreateActions()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(0, 4, 0, 8)
            };
        }
    }
}
