using System;
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

        /// <summary>创建 JDBC 环境窗体。</summary>
        /// <param name="store">环境设置存储。</param>
        /// <param name="settings">用于记住窗体尺寸的设置存储。</param>
        public JdbcEnvironmentForm(IJdbcEnvironmentStore store, ISettingsStore settings)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            Text = AppPresentation.WindowTitle("JDBC 环境");
            Width = 720;
            Height = 400;
            MinimumSize = new Size(560, 340);
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

            javaLayout.Controls.Add(CreateHint(
                "本页只配置运行 JDBC 所需的 Java（JDK 8 或更高）。请先点击「检测 Java」自动搜索本机；" +
                "若未找到，再点「打开 JDK 下载页」安装后重新检测。插件不附带 Java。"), 0, 0);
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

        private static Label CreateHint(string text)
        {
            return new Label
            {
                AutoSize = true,
                MaximumSize = new Size(1000, 0),
                Text = text,
                Padding = new Padding(0, 0, 0, 8)
            };
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
