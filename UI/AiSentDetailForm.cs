using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Models;
using DB2Sheet.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace DB2Sheet.UI
{
    /// <summary>只读显示一次请求的调用信息、占比条和分类原文。</summary>
    /// <remarks>
    /// 不记窗口位置。正文在 WebView2 里。调用信息始终可见，七个分类默认收起。
    /// 浏览器初始化失败时同一窗口改回纯文本，不打开外部浏览器。
    /// 必须在界面线程调用 <see cref="Show"/> 打开。
    /// </remarks>
    internal sealed class AiSentDetailForm : AppForm
    {
        private readonly WebView2 _view;
        private readonly TextBox _fallback;
        private readonly string _html;
        private readonly string _plain;
        private readonly string _browserData;
        private readonly ToolTip _tips = new ToolTip();
        private bool _browserStarted;

        /// <summary>摆 WebView2 和关闭按钮。页面在首次显示后加载。</summary>
        /// <param name="sections">这一轮的原文。</param>
        /// <param name="inputTokens">后一条助手回复的输入 token。没有时为空。</param>
        /// <param name="call">这次调用的地址和模型。旧对话为空。</param>
        /// <param name="browserData">与对话栏共用的 WebView2 用户数据目录。</param>
        private AiSentDetailForm(IList<AiSentSection> sections, int? inputTokens, AiSentCall call, string browserData)
        {
            Text = AppPresentation.WindowTitle("请求详情");
            Width = 640;
            Height = 1040;
            MinimumSize = new Size(420, 320);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            _html = AiSentReportHtml.Document(sections, inputTokens, call);
            _plain = AiSentReportHtml.Plain(sections, inputTokens, call);
            _browserData = browserData ?? string.Empty;

            _fallback = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                HideSelection = true,
                ScrollBars = ScrollBars.Both,
                BorderStyle = BorderStyle.None,
                WordWrap = true,
                Text = _plain,
                Visible = false
            };
            try
            {
                AiChatBrowserHost.PrepareLoader();
                _view = new WebView2 { Dock = DockStyle.Fill };
            }
            catch (Exception)
            {
                // 加载器缺失时改用纯文本，不让窗口因此打不开。
                _view = null;
                _fallback.Visible = true;
            }

            Panel host = new Panel { Dock = DockStyle.Fill };
            if (_view != null)
                host.Controls.Add(_view);
            host.Controls.Add(_fallback);

            Button close = new Button
            {
                Text = "关闭",
                AutoSize = true,
                DialogResult = DialogResult.OK,
                UseVisualStyleBackColor = true
            };
            _tips.SetToolTip(close, AppPresentation.AiSentCloseTip);
            FlowLayoutPanel actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(12, 8, 12, 12)
            };
            actions.Controls.Add(close);
            Controls.Add(host);
            Controls.Add(actions);
            AcceptButton = close;
            CancelButton = close;
            Shown += (sender, args) =>
            {
                close.Focus();
                BeginBrowser();
            };
        }

        /// <summary>打开只读窗。调用方必须在界面线程。</summary>
        /// <param name="owner">所属窗体。</param>
        /// <param name="sections">这一轮的原文。</param>
        /// <param name="inputTokens">后一条助手回复的输入 token。等待中或没返回时为空。</param>
        /// <param name="call">这次调用的地址和模型。旧对话为空。</param>
        /// <param name="browserData">与对话栏共用的 WebView2 用户数据目录。</param>
        public static void Show(IWin32Window owner, IList<AiSentSection> sections, int? inputTokens, AiSentCall call, string browserData)
        {
            using (AiSentDetailForm form = new AiSentDetailForm(sections, inputTokens, call, browserData))
                form.ShowDialog(owner);
        }

        /// <summary>释放悬浮提示。</summary>
        /// <param name="disposing">为 true 时释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _tips.Dispose();
            base.Dispose(disposing);
        }

        private async void BeginBrowser()
        {
            if (_browserStarted || _view == null || IsDisposed) return;
            _browserStarted = true;
            try
            {
                CoreWebView2Environment environment = await AiChatBrowserHost.CreateAsync(_browserData);
                if (IsDisposed) return;
                await _view.EnsureCoreWebView2Async(environment);
                if (IsDisposed || _view.CoreWebView2 == null) return;
                CoreWebView2Settings settings = _view.CoreWebView2.Settings;
                settings.AreDefaultScriptDialogsEnabled = false;
                settings.IsWebMessageEnabled = false;
                settings.AreHostObjectsAllowed = false;
                settings.AreDevToolsEnabled = false;
                settings.IsStatusBarEnabled = false;
                _view.CoreWebView2.NewWindowRequested += (sender, args) => { args.Handled = true; };
                _view.CoreWebView2.NavigationStarting += (sender, args) =>
                {
                    if (args.IsUserInitiated)
                        args.Cancel = true;
                };
                _view.CoreWebView2.NavigateToString(_html);
            }
            catch (Exception)
            {
                // 运行库缺失时改回纯文本。对话正文已经依赖 WebView2，这里不再弹一次错误详情。
                if (IsDisposed || _view == null) return;
                _view.Visible = false;
                _fallback.Visible = true;
            }
        }
    }
}
