using System;
using System.Collections.Generic;

namespace DB2Sheet.Models
{
    /// <summary>模型配置使用的请求协议。协议只决定报文格式，不代表某一家厂商。</summary>
    public enum AiProtocol
    {
        /// <summary>OpenAI 兼容的 chat/completions。多数厂商和网关走这一类。</summary>
        OpenAiCompatible,

        /// <summary>Anthropic 官方 Messages 接口。</summary>
        Claude,

        /// <summary>Google 官方 generateContent 接口。</summary>
        Gemini
    }

    /// <summary>一条用户保存的模型配置。密钥只以 DPAPI 密文出现。</summary>
    public sealed class AiModelProfile
    {
        /// <summary>稳定标识。新建时由存储分配。</summary>
        public string Id { get; set; }

        /// <summary>对话里显示的名称。</summary>
        public string DisplayName { get; set; }

        /// <summary>协议存储值：openai、claude 或 gemini。</summary>
        public string Protocol { get; set; }

        /// <summary>用户填写的接口地址，不含密钥。</summary>
        public string Endpoint { get; set; }

        /// <summary>旧文件里手填的单个型号。新配置改用 <see cref="ModelIds"/>，这里留空。</summary>
        public string ModelName { get; set; }

        /// <summary>测试通过时从接口读到的型号。对话里从这里选一个来调用。</summary>
        public List<string> ModelIds { get; set; }

        /// <summary>当前用户 DPAPI 保护后的密钥。明文不入库。</summary>
        public string ProtectedKey { get; set; }

        /// <summary>对话打开时优先选中的配置。同时只应有一条为 true。</summary>
        public bool IsDefault { get; set; }
    }

    /// <summary>一份只作为上下文的 Skill。正文是用户自己的说明，不会被执行。</summary>
    public sealed class AiSkill
    {
        /// <summary>稳定标识。</summary>
        public string Id { get; set; }

        /// <summary>对话里显示的名称。</summary>
        public string Name { get; set; }

        /// <summary>发给模型的上下文正文。</summary>
        public string Body { get; set; }

        /// <summary>上次保存时间，圆点往返格式。</summary>
        public string UpdatedUtc { get; set; }
    }

    /// <summary>对话中的一轮。只保留用户输入和助手回复，不保存模型密钥。</summary>
    public sealed class AiChatTurn
    {
        /// <summary>user 或 assistant。</summary>
        public string Role { get; set; }

        /// <summary>这一轮的可见文字。</summary>
        public string Text { get; set; }

        /// <summary>服务商返回的输入 token。只有助手轮次可能有值，缺了就不显示。</summary>
        public int? InputTokens { get; set; }

        /// <summary>服务商返回的输出 token。只有助手轮次可能有值，缺了就不显示。</summary>
        public int? OutputTokens { get; set; }

        /// <summary>这次发送前按类留存的原文。只有用户轮次可能有值，旧对话为空。</summary>
        public List<AiSentSection> Sent { get; set; }

        /// <summary>这次调用的地址、模型和响应。只有用户轮次可能有值，旧对话为空。不含密钥。</summary>
        public AiSentCall Call { get; set; }
    }

    /// <summary>查看请求详情时记下的这一次调用。只供本机查看，不进入审计日志。</summary>
    /// <remarks>查询串里的密钥参数在记下地址之前已经去掉。请求头和消息正文参数不保存。</remarks>
    public sealed class AiSentCall
    {
        /// <summary>实际 POST 的地址。</summary>
        public string Url { get; set; }

        /// <summary>协议存储值：openai、claude 或 gemini。</summary>
        public string Protocol { get; set; }

        /// <summary>发送当时的配置显示名。之后改名不影响这一轮。</summary>
        public string ModelDisplayName { get; set; }

        /// <summary>这次实际调用的子型号。</summary>
        public string ModelId { get; set; }

        /// <summary>发送当时选择的思考强度稳定值；旧对话为空。</summary>
        public string ThinkingLevel { get; set; }

        /// <summary>从实际请求体筛出的非内容参数 JSON；不含消息正文或认证信息。</summary>
        public string ApiParametersJson { get; set; }

        /// <summary>点击发送的时间，圆点往返格式的 UTC。</summary>
        public string SentUtc { get; set; }

        /// <summary>HTTP 状态码。还没返回时为空。</summary>
        public int? HttpStatus { get; set; }

        /// <summary>呼叫耗时，毫秒。还没返回时为空。</summary>
        public long? ElapsedMs { get; set; }
    }

    /// <summary>查看请求详情时的一段原文。正文是脱敏之前的文字，不含模型密钥。</summary>
    public sealed class AiSentSection
    {
        /// <summary>用户自己写的问题。</summary>
        public const string Question = "question";

        /// <summary>嵌入的系统说明，不含 Skill。</summary>
        public const string System = "system";

        /// <summary>一条 Skill 的名称和正文。</summary>
        public const string Skill = "skill";

        /// <summary>数据库类型说明。</summary>
        public const string Dialect = "dialect";

        /// <summary>当前查询页签的 SQL。</summary>
        public const string Tab = "tab";

        /// <summary>用 @ 附上的一条引用。</summary>
        public const string Citation = "citation";

        /// <summary>此前往来里看得见的问答。</summary>
        public const string History = "history";

        /// <summary>分类，取上面的常量。</summary>
        public string Kind { get; set; }

        /// <summary>这一段的名字。Skill 是名称，引用是芯片文字，页签是页签名。</summary>
        public string Title { get; set; }

        /// <summary>脱敏前的原文。</summary>
        public string Body { get; set; }
    }

    /// <summary>一次成功回复。用量缺一边时界面不显示计数。</summary>
    public sealed class AiReply
    {
        /// <summary>助手回复的文字。</summary>
        public string Text { get; set; }

        /// <summary>服务商返回的输入 token。没有该字段时为空。</summary>
        public int? InputTokens { get; set; }

        /// <summary>服务商返回的输出 token。没有该字段时为空。</summary>
        public int? OutputTokens { get; set; }

        /// <summary>HTTP 状态码。成功回复才有。</summary>
        public int? HttpStatus { get; set; }

        /// <summary>呼叫耗时，毫秒。成功回复才有。</summary>
        public long? ElapsedMs { get; set; }
    }

    /// <summary>把输入和输出 token 写成气泡下的一行。</summary>
    public static class AiTokenText
    {
        /// <summary>写成 ↑2,987 ↓234 tokens。千分位用逗号。</summary>
        /// <param name="input">输入 token。</param>
        /// <param name="output">输出 token。</param>
        /// <returns>灰色小字用的那一行。</returns>
        public static string Format(int input, int output)
        {
            return "↑" + Group(input) + " ↓" + Group(output) + " tokens";
        }

        private static string Group(int value)
        {
            return value.ToString("#,0", System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    /// <summary>一条可归档的本地对话。</summary>
    public sealed class AiConversation
    {
        /// <summary>稳定标识，也是文件名。</summary>
        public string Id { get; set; }

        /// <summary>用户看到的标题。</summary>
        public string Title { get; set; }

        /// <summary>用户改过名，或由报错入口指定标题时为 true。为 true 时不再用第一条问题覆盖标题。</summary>
        public bool TitleLocked { get; set; }

        /// <summary>归档后从当前列表收起，文件仍在。</summary>
        public bool Archived { get; set; }

        /// <summary>创建时间，圆点往返格式。</summary>
        public string CreatedUtc { get; set; }

        /// <summary>最后一次变更时间，圆点往返格式。</summary>
        public string UpdatedUtc { get; set; }

        /// <summary>按时间排列的往来。为空时表示还没有发送。</summary>
        public List<AiChatTurn> Messages { get; set; }
    }

    /// <summary>点击发送时从 SQL 编辑器读到的当前页签。</summary>
    public sealed class AiEditorSnapshot
    {
        /// <summary>有打开的页签时为 true。没有页签时 SQL 按空内容发送。</summary>
        public bool HasTab { get; set; }

        /// <summary>页签标题，不含未保存标记。</summary>
        public string TabName { get; set; }

        /// <summary>编辑器全文，含方案说明注释。</summary>
        public string Sql { get; set; }

        /// <summary>当前连接的提供程序标识。没有连接时为空。</summary>
        public string ProviderId { get; set; }
    }

    /// <summary>一次准备外发的审计内容。写入成功后才允许发送。</summary>
    public sealed class AiAuditEntry
    {
        /// <summary>动作名，例如 chat.send、model.test、conversation.delete。</summary>
        public string Action { get; set; }

        /// <summary>同一次调用的关联标识。待发记录和结果记录使用同一个值。</summary>
        public string OperationId { get; set; }

        /// <summary>pending 表示尚未发出；result 表示已经尝试呼叫。本地动作留空。</summary>
        public string Phase { get; set; }

        /// <summary>对话标识。与对话无关时为空。</summary>
        public string ConversationId { get; set; }

        /// <summary>对话标题。</summary>
        public string ConversationTitle { get; set; }

        /// <summary>模型配置的显示名。</summary>
        public string ModelDisplayName { get; set; }

        /// <summary>协议存储值。</summary>
        public string Protocol { get; set; }

        /// <summary>接口的主机、路径和查询参数名。不含查询参数值。</summary>
        public string Endpoint { get; set; }

        /// <summary>所选 Skill 名称。未选时为空。</summary>
        public string SkillName { get; set; }

        /// <summary>所选 Skill 全文。</summary>
        public string SkillBody { get; set; }

        /// <summary>实际放入请求的系统说明。</summary>
        public string SystemText { get; set; }

        /// <summary>用户这一次输入的全文。</summary>
        public string UserText { get; set; }

        /// <summary>此前往来，便于审计人员不必只靠哈希还原。</summary>
        public string HistoryText { get; set; }

        /// <summary>当前页签名称。</summary>
        public string TabName { get; set; }

        /// <summary>当前页签全文。没有页签时为空，并由 <see cref="HasTab"/> 标明。</summary>
        public string TabSql { get; set; }

        /// <summary>发送时是否有打开的页签。</summary>
        public bool? HasTab { get; set; }

        /// <summary>数据库类型说明。</summary>
        public string Dialect { get; set; }

        /// <summary>这次附上的查询方案、文件或预览。没有时为空。</summary>
        public string ContextText { get; set; }

        /// <summary>为 true 时这次发送包含预览结果。</summary>
        public bool? IncludesPreview { get; set; }

        /// <summary>脱敏时去掉了哪些内容。没有则为「无」。</summary>
        public string Redaction { get; set; }

        /// <summary>HTTP 请求体的 SHA256。</summary>
        public string RequestSha256 { get; set; }

        /// <summary>HTTP 状态码。还没发出或网络失败时为空或 0。</summary>
        public int? HttpStatus { get; set; }

        /// <summary>呼叫耗时，毫秒。</summary>
        public long? ElapsedMs { get; set; }

        /// <summary>模型回复，或失败时的响应正文。</summary>
        public string AssistantText { get; set; }

        /// <summary>本地动作的补充说明，例如是否更换了密钥。</summary>
        public string Detail { get; set; }
    }

    /// <summary>从配置里取出对话可选的型号。新列表优先，没有时用旧的单个型号名。</summary>
    public static class AiModelCatalog
    {
        /// <summary>返回去重后的型号。配置为空时得到空列表。</summary>
        /// <param name="profile">模型配置。</param>
        /// <returns>可在对话里选择的型号。</returns>
        public static IReadOnlyList<string> Ids(AiModelProfile profile)
        {
            List<string> ids = new List<string>();
            if (profile == null) return ids;
            if (profile.ModelIds != null)
            {
                foreach (string id in profile.ModelIds)
                    Add(ids, id);
            }

            if (ids.Count == 0)
                Add(ids, profile.ModelName);
            return ids;
        }

        private static void Add(List<string> ids, string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            string trimmed = id.Trim();
            foreach (string existing in ids)
            {
                if (string.Equals(existing, trimmed, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            ids.Add(trimmed);
        }
    }

    /// <summary>模型在设置列表、对话菜单和审计里的文字。</summary>
    public static class AiModelCaption
    {
        /// <summary>设置列表用的名称。默认配置在末尾加「（默认）」，不拼型号。</summary>
        /// <param name="profile">模型配置。</param>
        /// <returns>显示名。配置为空时为空字符串。</returns>
        public static string Format(AiModelProfile profile)
        {
            if (profile == null) return string.Empty;
            string text = profile.DisplayName ?? string.Empty;
            if (profile.IsDefault)
                text += "（默认）";
            return text;
        }

        /// <summary>对话下拉里的一行：显示名和子型号。</summary>
        /// <param name="displayName">配置显示名。</param>
        /// <param name="modelId">实际调用的型号。</param>
        /// <returns>中间空一格的文字。</returns>
        public static string Menu(string displayName, string modelId)
        {
            return (displayName ?? string.Empty) + " " + (modelId ?? string.Empty);
        }

        /// <summary>审计里的名称：显示名、斜杠、实际型号。</summary>
        /// <param name="displayName">配置显示名。</param>
        /// <param name="modelId">实际调用的型号。</param>
        /// <returns>可写进日志的一行。</returns>
        public static string Audit(string displayName, string modelId)
        {
            return (displayName ?? string.Empty) + " / " + (modelId ?? string.Empty);
        }
    }

    /// <summary>输入 @ 时能选到的已保存查询，以及当前页预览是否可附上。</summary>
    public sealed class AiMentionSource
    {
        /// <summary>已保存的查询方案。没有时为空列表。</summary>
        public IReadOnlyList<AiSavedQuery> Queries { get; set; }

        /// <summary>当前页签已经有预览结果时为 true。</summary>
        public bool HasPreview { get; set; }

        /// <summary>预览芯片上的名字，一般是页签标题。</summary>
        public string PreviewLabel { get; set; }

        /// <summary>已经截断过的预览文本。没有预览时为空。</summary>
        public string PreviewText { get; set; }
    }

    /// <summary>一条可被引用的已保存查询。</summary>
    public sealed class AiSavedQuery
    {
        /// <summary>方案名称。</summary>
        public string Name { get; set; }

        /// <summary>方案 SQL 全文。</summary>
        public string Sql { get; set; }
    }

    /// <summary>协议存储值和枚举之间的转换。</summary>
    public static class AiProtocolNames
    {
        /// <summary>OpenAI 兼容协议的存储值。</summary>
        public const string OpenAi = "openai";

        /// <summary>Claude 协议的存储值。</summary>
        public const string Claude = "claude";

        /// <summary>Gemini 协议的存储值。</summary>
        public const string Gemini = "gemini";

        /// <summary>把枚举写成存储值。</summary>
        /// <param name="protocol">协议。</param>
        /// <returns>openai、claude 或 gemini。</returns>
        public static string ToStorage(AiProtocol protocol)
        {
            switch (protocol)
            {
                case AiProtocol.Claude:
                    return Claude;
                case AiProtocol.Gemini:
                    return Gemini;
                default:
                    return OpenAi;
            }

        }

        /// <summary>从存储值读协议。无法识别时当作 OpenAI 兼容。</summary>
        /// <param name="storage">配置里的协议字符串。</param>
        /// <returns>对应协议。</returns>
        public static AiProtocol Parse(string storage)
        {
            if (string.Equals(storage, Claude, StringComparison.OrdinalIgnoreCase))
                return AiProtocol.Claude;
            if (string.Equals(storage, Gemini, StringComparison.OrdinalIgnoreCase))
                return AiProtocol.Gemini;
            return AiProtocol.OpenAiCompatible;
        }

        /// <summary>协议在窗体里的名称。</summary>
        /// <param name="protocol">协议。</param>
        /// <returns>用户可见名称。</returns>
        public static string DisplayName(AiProtocol protocol)
        {
            switch (protocol)
            {
                case AiProtocol.Claude:
                    return "Claude";
                case AiProtocol.Gemini:
                    return "Gemini";
                default:
                    return "OpenAI 兼容";
            }
        }
    }

    /// <summary>AI 思考强度选择的稳定会话值和界面文字。</summary>
    public static class AiThinkingLevels
    {
        /// <summary>不额外指定思考参数的默认档位。</summary>
        public const string Automatic = "automatic";
        /// <summary>低思考强度档位。</summary>
        public const string Low = "low";
        /// <summary>中思考强度档位。</summary>
        public const string Medium = "medium";
        /// <summary>高思考强度档位。</summary>
        public const string High = "high";

        /// <summary>把会话值转换为中文界面文字。</summary>
        /// <param name="value">automatic、low、medium 或 high。</param>
        /// <returns>自动、低、中或高。</returns>
        public static string DisplayName(string value)
        {
            if (string.Equals(value, Low, StringComparison.OrdinalIgnoreCase)) return "低";
            if (string.Equals(value, Medium, StringComparison.OrdinalIgnoreCase)) return "中";
            if (string.Equals(value, High, StringComparison.OrdinalIgnoreCase)) return "高";
            return "自动";
        }

        /// <summary>把未知或空的会话值归一为默认自动档位。</summary>
        /// <param name="value">设置文件中的档位值。</param>
        /// <returns>有效的稳定档位值。</returns>
        public static string Normalize(string value)
        {
            if (string.Equals(value, Low, StringComparison.OrdinalIgnoreCase)) return Low;
            if (string.Equals(value, Medium, StringComparison.OrdinalIgnoreCase)) return Medium;
            if (string.Equals(value, High, StringComparison.OrdinalIgnoreCase)) return High;
            return Automatic;
        }
    }
}
