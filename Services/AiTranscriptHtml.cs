using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text;
using DB2Sheet.Models;
using DB2Sheet.UI;
using Markdig;

namespace DB2Sheet.Services
{
    /// <summary>把一条对话画成不含模型脚本的 HTML，供 WebView2 显示。</summary>
    /// <remarks>
    /// 助手气泡在左、浅灰，用户气泡在右、浅蓝。用户原文只做转义。助手回复用 Markdown 排版，原始 HTML 不进入页面。
    /// 页面末尾有一段只负责滚到底部的脚本，模型正文里的脚本标签会变成文字。
    /// </remarks>
    public static class AiTranscriptHtml
    {
        private static readonly MarkdownPipeline Pipeline = new MarkdownPipelineBuilder()
            .UseAdvancedExtensions()
            .DisableHtml()
            .Build();
        private static readonly string CopyIcon = LoadCopyIcon();

        /// <summary>生成整页 HTML。</summary>
        /// <param name="messages">已经落盘的往来。可为空。</param>
        /// <param name="pendingUser">尚未落盘、正在等待回复的问题。为空表示没有。带发送原文时气泡下有链接。</param>
        /// <param name="waiting">为 true 时在刚发出的问题下面显示助手位的三个圆点。</param>
        /// <returns>完整 HTML 文档。</returns>
        public static string Document(IList<AiChatTurn> messages, AiChatTurn pendingUser, bool waiting)
        {
            StringBuilder body = new StringBuilder();
            if (messages != null)
            {
                for (int index = 0; index < messages.Count; index++)
                {
                    AiChatTurn turn = messages[index];
                    if (turn == null) continue;
                    AppendTurn(body, turn, index);
                }
            }

            if (pendingUser != null && !string.IsNullOrEmpty(pendingUser.Text))
                AppendTurn(body, pendingUser, -1);
            if (waiting)
            {
                body.Append("<div class=\"turn assistant\"><div class=\"bubble\">");
                body.Append("<div class=\"typing\"><span></span><span></span><span></span></div>");
                body.Append("</div></div>");
            }

            return "<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
                + "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data:; script-src 'unsafe-inline'\">"
                + "<style>"
                + "body{font-family:\"Microsoft YaHei UI\",sans-serif;font-size:13px;line-height:1.5;margin:12px;color:#222;background:#fff;}"
                + ".turn{display:flex;margin:10px 0;}"
                + ".turn.assistant{justify-content:flex-start;}"
                + ".turn.user{justify-content:flex-end;}"
                + ".stack{max-width:85%;min-width:0;}"
                + ".bubble{padding:8px 12px;border-radius:12px;box-sizing:border-box;}"
                + ".meta{display:block;margin-top:4px;font-size:11px;line-height:1.4;color:#888;text-decoration:none;}"
                + ".turn.user .meta{text-align:right;}"
                + "a.meta:hover{text-decoration:underline;}"
                + ".turn.assistant .bubble{background:#f3f3f3;}"
                + ".turn.user .bubble{background:#e7f0fe;white-space:pre-wrap;}"
                + ".bubble p{margin:0 0 8px;}"
                + ".bubble p:last-child{margin-bottom:0;}"
                + "pre,code{font-family:Consolas,monospace;font-size:12px;}"
                + "pre{padding:8px;overflow:auto;}"
                + "code{padding:0 3px;}"
                + ".codebox{position:relative;}"
                + ".codebox pre{padding-top:26px;}"
                + ".copy{position:absolute;top:4px;right:4px;border:0;background:transparent;padding:0;cursor:pointer;}"
                + ".copy img{width:16px;height:16px;display:block;}"
                + ".copied{display:none;position:absolute;top:2px;right:26px;font:12px \"Microsoft YaHei UI\",sans-serif;color:#666;pointer-events:none;white-space:nowrap;}"
                + ".turn.assistant pre,.turn.assistant code{background:#fff;}"
                + "pre code{background:transparent;padding:0;}"
                + "table{border-collapse:collapse;}"
                + "td,th{border:1px solid #ccc;padding:4px 8px;}"
                + "blockquote{border-left:3px solid #ccc;margin-left:0;padding-left:8px;color:#444;}"
                + ".typing span{display:inline-block;width:6px;height:6px;margin-right:4px;border-radius:50%;background:#888;animation:blink 1.2s infinite ease-in-out;}"
                + ".typing span:nth-child(2){animation-delay:.2s;}"
                + ".typing span:nth-child(3){animation-delay:.4s;}"
                + "@keyframes blink{0%,80%,100%{opacity:.2}40%{opacity:1}}"
                + "</style></head><body>"
                + body
                + CopyScript()
                + "</body></html>";
        }

        /// <summary>页面脚本：按下鼠标时通知宿主收起下拉；复制按钮成功或失败都在图标旁提示，约 1.5 秒后消失，不弹出窗体。</summary>
        /// <returns>含滚动到底部的脚本标签。</returns>
        private static string CopyScript()
        {
            return "<script>"
                + "document.addEventListener('pointerdown',function(){if(window.chrome&&chrome.webview)chrome.webview.postMessage('pointerdown');},true);"
                + "document.addEventListener('click',function(e){"
                + "var node=e.target;while(node&&(!node.classList||!node.classList.contains('copy')))node=node.parentNode;"
                + "if(!node)return;var box=node.parentNode;var pre=box?box.querySelector('pre'):null;if(!pre)return;"
                + "var area=document.createElement('textarea');area.value=pre.innerText;"
                + "document.body.appendChild(area);area.select();var copied=false;"
                + "try{copied=document.execCommand('copy');}catch(err){copied=false;}"
                + "document.body.removeChild(area);"
                + "var tip=box.querySelector('.copied');"
                + "if(!tip){tip=document.createElement('span');tip.className='copied';box.insertBefore(tip,node);}"
                + "tip.textContent=copied?'" + Js(AppPresentation.AiCodeCopied) + "':'" + Js(AppPresentation.AiCodeCopyFailed) + "';"
                + "tip.style.display='block';"
                + "if(tip._hide)clearTimeout(tip._hide);"
                + "tip._hide=setTimeout(function(){tip.style.display='none';},1500);"
                + "});"
                + "window.scrollTo(0,document.body.scrollHeight);</script>";
        }

        /// <summary>放进脚本字符串。只转义会截断引号的字符。</summary>
        /// <param name="text">界面文案。</param>
        /// <returns>可夹在单引号里的文字。</returns>
        private static string Js(string text)
        {
            return (text ?? string.Empty).Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", string.Empty).Replace("\n", "\\n");
        }

        /// <summary>把助手回复转成 HTML 片段。供测试核对表格、代码和脚本转义。</summary>
        /// <param name="markdown">模型返回的原文。</param>
        /// <returns>片段。脚本标签以文字形式出现。</returns>
        public static string AssistantFragment(string markdown)
        {
            return Markdown.ToHtml(markdown ?? string.Empty, Pipeline);
        }

        private static string DecorateCode(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;
            StringBuilder builder = new StringBuilder();
            int cursor = 0;
            while (cursor < html.Length)
            {
                int open = html.IndexOf("<pre", cursor, StringComparison.OrdinalIgnoreCase);
                if (open < 0)
                {
                    builder.Append(html, cursor, html.Length - cursor);
                    break;
                }

                int openEnd = html.IndexOf('>', open);
                int close = openEnd < 0 ? -1 : html.IndexOf("</pre>", openEnd, StringComparison.OrdinalIgnoreCase);
                if (openEnd < 0 || close < 0)
                {
                    builder.Append(html, cursor, html.Length - cursor);
                    break;
                }

                builder.Append(html, cursor, open - cursor);
                builder.Append("<div class=\"codebox\"><button type=\"button\" class=\"copy\" title=\"复制\">");
                if (CopyIcon.Length > 0)
                    builder.Append("<img alt=\"复制\" src=\"data:image/png;base64,").Append(CopyIcon).Append("\">");
                else
                    builder.Append("复制");
                builder.Append("</button>");
                builder.Append(html, open, close + "</pre>".Length - open);
                builder.Append("</div>");
                cursor = close + "</pre>".Length;
            }

            return builder.ToString();
        }

        private static string LoadCopyIcon()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("DB2Sheet.Resources.icon_copy.png"))
            {
                if (stream == null) return string.Empty;
                using (MemoryStream memory = new MemoryStream())
                {
                    stream.CopyTo(memory);
                    return Convert.ToBase64String(memory.ToArray());
                }
            }
        }

        private static void AppendTurn(StringBuilder body, AiChatTurn turn, int index)
        {
            bool assistant = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase);
            body.Append(assistant ? "<div class=\"turn assistant\">" : "<div class=\"turn user\">");
            body.Append("<div class=\"stack\"><div class=\"bubble\">");
            if (assistant)
                body.Append(DecorateCode(AssistantFragment(turn.Text)));
            else
                body.Append(WebUtility.HtmlEncode(turn.Text ?? string.Empty));
            body.Append("</div>");
            if (!assistant && turn.Sent != null && turn.Sent.Count > 0)
            {
                string href = index < 0
                    ? "https://db2sheet.invalid/sent/pending"
                    : "https://db2sheet.invalid/sent/" + index.ToString(System.Globalization.CultureInfo.InvariantCulture);
                body.Append("<a class=\"meta\" href=\"").Append(href).Append("\" title=\"");
                body.Append(TipAttribute(AppPresentation.AiSentLinkTip));
                body.Append("\">").Append(WebUtility.HtmlEncode(AppPresentation.AiSentLink)).Append("</a>");
            }

            if (assistant && turn.InputTokens.HasValue && turn.OutputTokens.HasValue)
            {
                body.Append("<div class=\"meta\">");
                body.Append(WebUtility.HtmlEncode(AiTokenText.Format(turn.InputTokens.Value, turn.OutputTokens.Value)));
                body.Append("</div>");
            }

            body.Append("</div></div>");
        }

        private static string TipAttribute(string tip)
        {
            string encoded = WebUtility.HtmlEncode(tip ?? string.Empty);
            return encoded.Replace("\r\n", "&#10;").Replace("\n", "&#10;");
        }
    }
}
