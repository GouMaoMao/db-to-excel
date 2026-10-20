using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;
using DB2Sheet.Storage;

namespace DB2Sheet.Services
{
    /// <summary>组装模型、Skill、对话、审计和三种协议的提问。</summary>
    /// <remarks>
    /// 呼叫模型时先写审计。审计写失败不会发出请求。
    /// 不读取数据库，也不执行返回的 SQL。
    /// </remarks>
    public sealed class AiAssistant
    {
        private readonly AiChatClient _client = new AiChatClient();

        /// <summary>创建助手并打开本机存储。</summary>
        /// <param name="paths">本地数据目录。</param>
        /// <param name="logger">日常索引日志。</param>
        public AiAssistant(ApplicationPaths paths, ILogger logger)
        {
            if (paths == null) throw new ArgumentNullException(nameof(paths));
            Models = new AiModelStore(paths);
            Skills = new AiSkillStore(paths);
            Conversations = new AiConversationStore(paths);
            Audit = new AiAuditLog(paths, logger);
            BrowserDataDirectory = paths.WebView2UserDataDirectory;
        }

        /// <summary>模型配置存储。</summary>
        public AiModelStore Models { get; }

        /// <summary>Skill 存储。</summary>
        public AiSkillStore Skills { get; }

        /// <summary>对话存储。</summary>
        public AiConversationStore Conversations { get; }

        /// <summary>审计日志。</summary>
        public AiAuditLog Audit { get; }

        /// <summary>WebView2 用户数据目录。对话页的浏览器缓存写在这里。</summary>
        public string BrowserDataDirectory { get; }

        /// <summary>用当前输入做一次提问。测试连接时不带 Skill、页签和系统说明。</summary>
        /// <param name="profile">选中的模型配置。</param>
        /// <param name="apiKey">已经解开的密钥。</param>
        /// <param name="modelName">这次实际调用的型号。</param>
        /// <param name="entry">审计字段。方法会补上哈希、状态和回复。</param>
        /// <param name="history">此前往来。测试时为空。</param>
        /// <param name="systemText">系统说明。测试时为空。</param>
        /// <param name="finalUserText">这一次发给模型的用户消息。</param>
        /// <returns>助手回复、可选的 token 计数，以及这次的 HTTP 状态和耗时。</returns>
        /// <exception cref="InvalidOperationException">地址、响应或结果审计失败。</exception>
        /// <remarks>
        /// 发出前的审计在调用方线程上完成。进入等待后，网络、解析和结果审计不回到该线程。
        /// 调用方只在拿到回复后更新界面，后台线程不访问 Excel。
        /// </remarks>
        public async Task<AiReply> AskAsync(
            AiModelProfile profile,
            string apiKey,
            string modelName,
            AiAuditEntry entry,
            IList<AiChatTurn> history,
            string systemText,
            string finalUserText)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            if (string.IsNullOrWhiteSpace(modelName))
                throw new InvalidOperationException("请先选择型号。");

            List<AiChatTurn> turns = new List<AiChatTurn>();
            if (history != null)
                turns.AddRange(history.Where(item => item != null));
            turns.Add(new AiChatTurn { Role = "user", Text = finalUserText ?? string.Empty });

            AiProtocol protocol = AiProtocolNames.Parse(profile.Protocol);
            AiChatClient.OutboundCall call = _client.Build(
                protocol,
                profile.Endpoint,
                apiKey,
                modelName.Trim(),
                systemText,
                turns);
            entry.OperationId = string.IsNullOrWhiteSpace(entry.OperationId) ? Guid.NewGuid().ToString("N") : entry.OperationId;
            entry.RequestSha256 = AiAuditLog.Sha256(call.Body);
            entry.Phase = "pending";
            entry.HttpStatus = null;
            entry.ElapsedMs = null;
            entry.AssistantText = null;
            Audit.Write(entry);

            Stopwatch watch = Stopwatch.StartNew();
            AiChatClient.HttpResult result;
            try
            {
                result = await _client.SendAsync(call).ConfigureAwait(false);
            }
            catch (Exception exception)
            {
                watch.Stop();
                entry.Phase = "result";
                entry.ElapsedMs = watch.ElapsedMilliseconds;
                entry.HttpStatus = 0;
                entry.AssistantText = AiTextRedactor.Redact(exception.Message, out string _);
                Audit.Write(entry);
                throw;
            }

            watch.Stop();
            string redactedBody = AiTextRedactor.Redact(result.Body, out string _);
            entry.Phase = "result";
            entry.ElapsedMs = watch.ElapsedMilliseconds;
            entry.HttpStatus = result.StatusCode;
            if (result.StatusCode < 200 || result.StatusCode >= 300)
            {
                entry.AssistantText = AiTextRedactor.Redact(_client.ReadError(redactedBody), out string _);
                Audit.Write(entry);
                throw new InvalidOperationException(
                    "模型请求失败（HTTP " + result.StatusCode.ToString() + "）。" + entry.AssistantText);
            }

            string reply;
            try
            {
                reply = _client.ReadReply(protocol, result.Body);
            }
            catch (Exception exception)
            {
                entry.AssistantText = redactedBody;
                Audit.Write(entry);
                throw new InvalidOperationException(exception.Message, exception);
            }

            entry.AssistantText = reply;
            AiReply usage = _client.ReadUsage(protocol, result.Body);
            try
            {
                Audit.Write(entry);
            }
            catch (Exception logException)
            {
                throw new IOException(
                    "模型已返回，但审计日志写入失败。回复如下：" + Environment.NewLine + reply,
                    logException);
            }

            return new AiReply
            {
                Text = reply,
                InputTokens = usage.InputTokens,
                OutputTokens = usage.OutputTokens,
                HttpStatus = result.StatusCode,
                ElapsedMs = watch.ElapsedMilliseconds
            };
        }

        /// <summary>向厂商拉取型号列表。</summary>
        /// <param name="protocol">协议。</param>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <param name="apiKey">明文密钥。</param>
        /// <returns>型号列表。</returns>
        /// <remarks>不写审计。测试窗体在随后的探测对话里写。</remarks>
        public Task<IReadOnlyList<string>> ListModelsAsync(AiProtocol protocol, string endpoint, string apiKey)
        {
            return _client.ListModelsAsync(protocol, endpoint, apiKey);
        }

        /// <summary>从列表里挑一个用来探测对话的型号。优先像聊天模型的 id，否则用第一项。</summary>
        /// <param name="ids">接口返回的型号。</param>
        /// <returns>探测用的型号。</returns>
        public static string ChooseProbeModel(IReadOnlyList<string> ids)
        {
            if (ids == null || ids.Count == 0)
                throw new InvalidOperationException("接口没有返回型号。");
            string[] hints = { "chat", "gpt", "claude", "gemini", "deepseek", "qwen", "glm", "flash" };
            foreach (string id in ids)
            {
                string lower = (id ?? string.Empty).ToLowerInvariant();
                foreach (string hint in hints)
                {
                    if (lower.Contains(hint))
                        return id;
                }
            }

            return ids[0];
        }

        /// <summary>记下一条不含正文以外字段的本地动作。</summary>
        /// <param name="action">动作名。</param>
        /// <param name="detail">补充说明。</param>
        /// <param name="conversationId">对话标识，可为空。</param>
        /// <param name="conversationTitle">对话标题，可为空。</param>
        public void WriteAction(string action, string detail, string conversationId, string conversationTitle)
        {
            Audit.Write(new AiAuditEntry
            {
                Action = action,
                Detail = detail,
                ConversationId = conversationId,
                ConversationTitle = conversationTitle,
                OperationId = Guid.NewGuid().ToString("N")
            });
        }

        /// <summary>拼出这一次用户消息，带上数据库类型和当前页签。正文来自嵌入的提示词。</summary>
        /// <param name="dialect">数据库类型说明。</param>
        /// <param name="hasTab">是否有打开的页签。</param>
        /// <param name="tabName">页签名称。</param>
        /// <param name="tabSql">已经脱敏的页签全文。</param>
        /// <param name="userText">已经脱敏的用户输入。</param>
        /// <param name="context">已经脱敏的额外引用。没有时为空。</param>
        /// <returns>发给模型的用户消息。</returns>
        public static string ComposeUserMessage(string dialect, bool hasTab, string tabName, string tabSql, string userText, string context)
        {
            return AiPrompt.User(dialect, hasTab, tabName, tabSql, userText, context);
        }

        /// <summary>把 Skill 接到系统说明后面。没有 Skill 时只返回已填入子型号的系统说明。</summary>
        /// <param name="modelId">对话里选中的子型号。</param>
        /// <param name="skillName">Skill 名称。</param>
        /// <param name="skillBody">Skill 正文。</param>
        /// <returns>请求里的系统说明。子型号已经写入，不会留下未替换的占位。</returns>
        public static string ComposeSystem(string modelId, string skillName, string skillBody)
        {
            string system = AiPrompt.SystemText(modelId);
            if (string.IsNullOrWhiteSpace(skillBody)) return system;
            return system + "\r\n\r\n" + AiPrompt.Skill(skillName, skillBody);
        }

        /// <summary>把此前往来收成审计可读的文本。</summary>
        /// <param name="history">已保存的往来。</param>
        /// <returns>多行文本。没有往来时为空。</returns>
        public static string FormatHistory(IList<AiChatTurn> history)
        {
            if (history == null || history.Count == 0) return string.Empty;
            StringBuilder builder = new StringBuilder();
            foreach (AiChatTurn turn in history)
            {
                if (turn == null) continue;
                builder.Append(turn.Role ?? "user").AppendLine(":");
                builder.AppendLine(turn.Text ?? string.Empty);
            }

            return builder.ToString();
        }

        /// <summary>从末尾截取看得见的问答，供生成标题。不含发送快照。</summary>
        /// <param name="messages">已保存的往来。</param>
        /// <param name="maxChars">最多保留的字符数。</param>
        /// <returns>近期问答。没有往来时为空。</returns>
        /// <remarks>标题请求单独计费，所以只带截断后的可见文字，不带 SQL、Skill 和引用。</remarks>
        public static string RecentVisible(IList<AiChatTurn> messages, int maxChars)
        {
            if (messages == null || messages.Count == 0 || maxChars <= 0) return string.Empty;
            List<string> parts = new List<string>();
            int used = 0;
            for (int index = messages.Count - 1; index >= 0; index--)
            {
                AiChatTurn turn = messages[index];
                if (turn == null) continue;
                string role = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "助手" : "用户";
                string block = role + "：\r\n" + (turn.Text ?? string.Empty);
                if (parts.Count > 0 && used + block.Length > maxChars)
                    break;
                if (block.Length > maxChars)
                    block = block.Substring(block.Length - maxChars);
                parts.Add(block);
                used += block.Length;
                if (used >= maxChars) break;
            }

            parts.Reverse();
            return string.Join("\r\n\r\n", parts);
        }

    }
}
