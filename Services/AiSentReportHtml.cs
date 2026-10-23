using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;
using System.Text;
using DB2Sheet.Models;
using DB2Sheet.UI;

namespace DB2Sheet.Services
{
    /// <summary>把一次发送画成调用信息、占比条和默认收起的分类原文。</summary>
    /// <remarks>
    /// 服务商不返回每一段的 token。有整次输入 token 时按字数占比拆成整数，各段相加等于该数。
    /// 没有时用字数估算：汉字计 1，英文约 4 个字母计 1。页面不含脚本，原文只转义。
    /// 调用信息里的地址已经去掉密钥查询参数，不放请求头。
    /// </remarks>
    public static class AiSentReportHtml
    {
        private const string SystemColor = "#9aa0a6";
        private const string SkillColor = "#7b61c9";
        private const string DialectColor = "#3d9a5f";
        private const string TabColor = "#d4a017";
        private const string CitationColor = "#5c4d9a";
        private const string HistoryColor = "#3b6fd8";
        private const string QuestionColor = "#e07a2f";

        /// <summary>生成整页 HTML。</summary>
        /// <param name="sections">脱敏前的分类原文。可为空。</param>
        /// <param name="inputTokens">紧随这条用户消息的助手回复给出的输入 token。没有时为空。</param>
        /// <param name="call">这次调用的地址和模型。旧对话为空。</param>
        /// <returns>完整 HTML 文档。</returns>
        public static string Document(IList<AiSentSection> sections, int? inputTokens, AiSentCall call)
        {
            List<Slice> slices = Build(sections, inputTokens);
            StringBuilder body = new StringBuilder();
            AppendCall(body, call);
            body.Append("<div class=\"bar\">");
            foreach (Slice slice in slices)
            {
                if (slice.Tokens <= 0) continue;
                body.Append("<i style=\"flex:").Append(slice.Tokens.ToString(CultureInfo.InvariantCulture));
                body.Append(" 1 0;background:").Append(slice.Color).Append("\"></i>");
            }

            body.Append("</div><p class=\"note\">").Append(WebUtility.HtmlEncode(Note(inputTokens))).Append("</p>");
            foreach (Slice slice in slices)
                AppendDetails(body, slice);
            return "<!DOCTYPE html><html><head><meta charset=\"utf-8\">"
                + "<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'\">"
                + "<style>"
                + "body{font-family:\"Microsoft YaHei UI\",sans-serif;font-size:13px;line-height:1.5;margin:16px;color:#222;background:#fff;}"
                + ".call{margin:0 0 16px;}"
                + ".call div{display:flex;gap:12px;margin:0 0 6px;}"
                + ".call .k{color:#888;flex:0 0 4.5em;}"
                + ".call .v{flex:1;word-break:break-all;}"
                + ".bar{display:flex;height:10px;border-radius:5px;overflow:hidden;background:#eee;}"
                + ".bar i{display:block;height:10px;min-width:2px;}"
                + ".note{color:#888;font-size:12px;margin:8px 0 16px;}"
                + "details{border-bottom:1px solid #eee;}"
                + "summary{display:flex;align-items:center;gap:8px;padding:8px 0;cursor:pointer;}"
                + "summary::-webkit-details-marker{display:none;}"
                + ".swatch{width:10px;height:10px;border-radius:2px;flex:0 0 auto;}"
                + ".name{flex:1;}"
                + ".count{color:#666;}"
                + ".preface{color:#888;font-size:12px;margin:0 0 8px;}"
                + ".piece-title{font-weight:600;margin:8px 0 4px;}"
                + "pre{white-space:pre-wrap;word-break:break-word;font-family:Consolas,\"Microsoft YaHei UI\",sans-serif;font-size:12px;background:#f6f6f6;padding:8px;margin:0 0 12px;}"
                + "</style></head><body>"
                + body
                + "</body></html>";
        }

        /// <summary>WebView2 不可用时的纯文本。分段和条形图说明与 HTML 一致。</summary>
        /// <param name="sections">脱敏前的分类原文。可为空。</param>
        /// <param name="inputTokens">紧随这条用户消息的助手回复给出的输入 token。没有时为空。</param>
        /// <param name="call">这次调用的地址和模型。旧对话为空。</param>
        /// <returns>多行文本。</returns>
        public static string Plain(IList<AiSentSection> sections, int? inputTokens, AiSentCall call)
        {
            List<Slice> slices = Build(sections, inputTokens);
            StringBuilder text = new StringBuilder();
            AppendCallPlain(text, call);
            text.AppendLine(Note(inputTokens));
            foreach (Slice slice in slices)
            {
                text.Append(slice.Title).Append("  ").AppendLine(CountText(slice));
                if (!string.IsNullOrEmpty(slice.Preface))
                    text.AppendLine(slice.Preface);
                AppendPlainPieces(text, slice);
                text.AppendLine();
            }

            return text.ToString().TrimEnd();
        }

        /// <summary>色条上方的调用信息。没有记录时只留一句说明。</summary>
        private static void AppendCall(StringBuilder body, AiSentCall call)
        {
            if (call == null)
            {
                body.Append("<p class=\"note\">").Append(WebUtility.HtmlEncode(AppPresentation.AiSentMissingCall)).Append("</p>");
                return;
            }

            body.Append("<div class=\"call\">");
            AppendFact(body, AppPresentation.AiSentUrl, PostUrl(call));
            AppendFact(body, AppPresentation.AiSentProtocol, ProtocolName(call));
            AppendFact(body, AppPresentation.AiSentModel, AiModelCaption.Audit(call.ModelDisplayName, call.ModelId));
            AppendFact(body, AppPresentation.AiSentThinking, ThinkingLevel(call));
            AppendFact(body, AppPresentation.AiSentTime, FormatSent(call.SentUtc));
            AppendFact(body, AppPresentation.AiSentResponse, FormatResponse(call));
            body.Append("<div class=\"api-parameters\"><div class=\"api-title\">")
                .Append(WebUtility.HtmlEncode(AppPresentation.AiSentApiParameters))
                .Append("</div><pre>");
            body.Append(WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(call.ApiParametersJson)
                ? AppPresentation.AiSentNoApiParameters
                : call.ApiParametersJson));
            body.Append("</pre></div>");
            body.Append("</div>");
        }

        /// <summary>调用信息里的一行。文字先转义。</summary>
        private static void AppendFact(StringBuilder body, string label, string value)
        {
            body.Append("<div><span class=\"k\">").Append(WebUtility.HtmlEncode(label)).Append("</span>");
            body.Append("<span class=\"v\">").Append(WebUtility.HtmlEncode(value)).Append("</span></div>");
        }

        /// <summary>纯文本里的调用信息，行序与页面一致。</summary>
        private static void AppendCallPlain(StringBuilder text, AiSentCall call)
        {
            if (call == null)
            {
                text.AppendLine(AppPresentation.AiSentMissingCall);
                text.AppendLine();
                return;
            }

            text.Append(AppPresentation.AiSentUrl).Append("  ").AppendLine(PostUrl(call));
            text.Append(AppPresentation.AiSentProtocol).Append("  ").AppendLine(ProtocolName(call));
            text.Append(AppPresentation.AiSentModel).Append("  ").AppendLine(AiModelCaption.Audit(call.ModelDisplayName, call.ModelId));
            text.Append(AppPresentation.AiSentThinking).Append("  ").AppendLine(ThinkingLevel(call));
            text.Append(AppPresentation.AiSentTime).Append("  ").AppendLine(FormatSent(call.SentUtc));
            text.Append(AppPresentation.AiSentResponse).Append("  ").AppendLine(FormatResponse(call));
            text.Append(AppPresentation.AiSentApiParameters).AppendLine(":");
            text.AppendLine(string.IsNullOrWhiteSpace(call.ApiParametersJson)
                ? AppPresentation.AiSentNoApiParameters
                : call.ApiParametersJson);
            text.AppendLine();
        }

        private static string ThinkingLevel(AiSentCall call)
        {
            return string.IsNullOrWhiteSpace(call.ThinkingLevel)
                ? AppPresentation.AiSentEmpty
                : AiThinkingLevels.DisplayName(call.ThinkingLevel);
        }

        /// <summary>地址前面加上 POST，标明这次发出的方法。</summary>
        private static string PostUrl(AiSentCall call)
        {
            return string.Format(CultureInfo.InvariantCulture, AppPresentation.AiSentPostUrl, call.Url ?? string.Empty);
        }

        /// <summary>把协议存储值写成窗体里的名称。</summary>
        private static string ProtocolName(AiSentCall call)
        {
            return AiProtocolNames.DisplayName(AiProtocolNames.Parse(call.Protocol));
        }

        /// <summary>把 UTC 圆点时间显示成本地的年月日时分秒。解析不了时原样显示。</summary>
        private static string FormatSent(string sentUtc)
        {
            if (string.IsNullOrWhiteSpace(sentUtc)) return AppPresentation.AiSentEmpty;
            DateTime parsed;
            if (!DateTime.TryParse(sentUtc, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out parsed))
                return sentUtc;
            return parsed.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        /// <summary>还没返回时写尚未返回。有状态码时写 HTTP 状态和毫秒耗时。</summary>
        private static string FormatResponse(AiSentCall call)
        {
            if (call == null || !call.HttpStatus.HasValue)
                return AppPresentation.AiSentPendingResponse;
            long elapsed = call.ElapsedMs.HasValue && call.ElapsedMs.Value > 0 ? call.ElapsedMs.Value : 0;
            return string.Format(
                CultureInfo.InvariantCulture,
                AppPresentation.AiSentResponseValue,
                call.HttpStatus.Value.ToString(CultureInfo.InvariantCulture),
                elapsed.ToString("#,0", CultureInfo.InvariantCulture));
        }

        private static void AppendDetails(StringBuilder body, Slice slice)
        {
            body.Append("<details><summary><i class=\"swatch\" style=\"background:").Append(slice.Color).Append("\"></i>");
            body.Append("<span class=\"name\">").Append(WebUtility.HtmlEncode(slice.Title)).Append("</span>");
            body.Append("<span class=\"count\">").Append(WebUtility.HtmlEncode(CountText(slice))).Append("</span></summary>");
            if (!string.IsNullOrEmpty(slice.Preface))
                body.Append("<p class=\"preface\">").Append(WebUtility.HtmlEncode(slice.Preface)).Append("</p>");
            if (slice.Pieces.Count == 0)
                body.Append("<pre>").Append(WebUtility.HtmlEncode(AppPresentation.AiSentEmpty)).Append("</pre>");
            else
            {
                foreach (Piece piece in slice.Pieces)
                {
                    if (slice.Pieces.Count > 1 || !string.IsNullOrWhiteSpace(piece.Title))
                        body.Append("<div class=\"piece-title\">").Append(WebUtility.HtmlEncode(PieceHeading(piece))).Append("</div>");
                    body.Append("<pre>").Append(WebUtility.HtmlEncode(DisplayBody(piece.Body))).Append("</pre>");
                }
            }

            body.Append("</details>");
        }

        private static void AppendPlainPieces(StringBuilder text, Slice slice)
        {
            if (slice.Pieces.Count == 0)
            {
                text.AppendLine(AppPresentation.AiSentEmpty);
                return;
            }

            foreach (Piece piece in slice.Pieces)
            {
                if (slice.Pieces.Count > 1 || !string.IsNullOrWhiteSpace(piece.Title))
                    text.AppendLine(PieceHeading(piece));
                text.AppendLine(DisplayBody(piece.Body));
            }
        }

        private static string Note(int? inputTokens)
        {
            if (!inputTokens.HasValue)
                return AppPresentation.AiSentEstimateNote;
            int total = inputTokens.Value < 0 ? 0 : inputTokens.Value;
            return string.Format(
                CultureInfo.InvariantCulture,
                AppPresentation.AiSentShareNote,
                total.ToString("#,0", CultureInfo.InvariantCulture));
        }

        private static string CountText(Slice slice)
        {
            return slice.Weight == 0 ? AppPresentation.AiSentEmpty : Compact(slice.Tokens);
        }

        private static string PieceHeading(Piece piece)
        {
            return string.IsNullOrWhiteSpace(piece.Title) ? AppPresentation.AiSentEmpty : piece.Title.Trim();
        }

        private static string DisplayBody(string body)
        {
            return string.IsNullOrWhiteSpace(body) ? AppPresentation.AiSentEmpty : body;
        }

        /// <summary>一千及以上写成 1.7K，更小的写成整数。</summary>
        private static string Compact(int value)
        {
            if (value < 0) value = 0;
            if (value < 1000) return value.ToString(CultureInfo.InvariantCulture);
            return (value / 1000.0).ToString("0.#", CultureInfo.InvariantCulture) + "K";
        }

        private static List<Slice> Build(IList<AiSentSection> sections, int? inputTokens)
        {
            List<Slice> slices = new List<Slice>
            {
                Make(AppPresentation.AiSentSystem, SystemColor, OfKind(sections, AiSentSection.System), null),
                Make(AppPresentation.AiSentSkill, SkillColor, OfKind(sections, AiSentSection.Skill), null),
                Make(AppPresentation.AiSentDialect, DialectColor, OfKind(sections, AiSentSection.Dialect), null),
                Make(AppPresentation.AiSentTab, TabColor, OfKind(sections, AiSentSection.Tab), null),
                Make(AppPresentation.AiSentCitation, CitationColor, OfKind(sections, AiSentSection.Citation), null),
                Make(AppPresentation.AiSentHistory, HistoryColor, OfKind(sections, AiSentSection.History), AppPresentation.AiSentHistoryNote),
                Make(AppPresentation.AiSentQuestion, QuestionColor, OfKind(sections, AiSentSection.Question), null)
            };
            AssignTokens(slices, inputTokens);
            return slices;
        }

        private static Slice Make(string title, string color, List<Piece> pieces, string preface)
        {
            Slice slice = new Slice
            {
                Title = title,
                Color = color,
                Preface = preface,
                Pieces = pieces
            };
            int weight = 0;
            foreach (Piece piece in pieces)
                weight += Estimate(piece.Body);
            slice.Weight = weight;
            return slice;
        }

        private static List<Piece> OfKind(IList<AiSentSection> sections, string kind)
        {
            List<Piece> pieces = new List<Piece>();
            if (sections == null) return pieces;
            foreach (AiSentSection section in sections)
            {
                if (section == null || !string.Equals(section.Kind, kind, StringComparison.Ordinal)) continue;
                pieces.Add(new Piece { Title = section.Title ?? string.Empty, Body = section.Body ?? string.Empty });
            }

            return pieces;
        }

        /// <summary>汉字计 1，其余非空白约 4 个计 1。空白不计。</summary>
        private static int Estimate(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            int ascii = 0;
            int tokens = 0;
            for (int index = 0; index < text.Length; index++)
            {
                char current = text[index];
                if (char.IsWhiteSpace(current)) continue;
                if (current <= 127) ascii++;
                else tokens++;
            }

            return tokens + (ascii + 3) / 4;
        }

        private static void AssignTokens(List<Slice> slices, int? inputTokens)
        {
            int weight = 0;
            foreach (Slice slice in slices)
                weight += slice.Weight;
            if (!inputTokens.HasValue)
            {
                foreach (Slice slice in slices)
                    slice.Tokens = slice.Weight;
                return;
            }

            int total = inputTokens.Value < 0 ? 0 : inputTokens.Value;
            if (weight == 0 || total == 0)
            {
                foreach (Slice slice in slices)
                    slice.Tokens = 0;
                return;
            }

            double[] fraction = new double[slices.Count];
            int assigned = 0;
            for (int index = 0; index < slices.Count; index++)
            {
                double exact = (double)slices[index].Weight * total / weight;
                int floor = (int)Math.Floor(exact);
                slices[index].Tokens = floor;
                fraction[index] = exact - floor;
                assigned += floor;
            }

            int left = total - assigned;
            int cursor = 0;
            while (left > 0)
            {
                int best = -1;
                for (int index = 0; index < slices.Count; index++)
                {
                    if (fraction[index] <= 0) continue;
                    if (best < 0 || fraction[index] > fraction[best])
                        best = index;
                }

                if (best < 0)
                {
                    slices[cursor % slices.Count].Tokens++;
                    cursor++;
                }
                else
                {
                    slices[best].Tokens++;
                    fraction[best] = 0;
                }

                left--;
            }
        }

        private sealed class Slice
        {
            public string Title;
            public string Color;
            public string Preface;
            public List<Piece> Pieces;
            public int Weight;
            public int Tokens;
        }

        private sealed class Piece
        {
            public string Title;
            public string Body;
        }
    }
}
