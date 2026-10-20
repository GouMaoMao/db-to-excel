using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>一条对话一个 JSON 文件。删除文件不会触碰审计日志。</summary>
    public sealed class AiConversationStore
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        private List<AiConversation> _conversations;

        /// <summary>创建存储并读取已有对话。</summary>
        /// <param name="paths">本地数据目录。</param>
        public AiConversationStore(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _conversations = Load();
        }

        /// <summary>对话变化后触发。在锁外调用，查询窗和弹出窗都靠它刷新列表。</summary>
        public event EventHandler Changed;

        /// <summary>返回全部对话副本。当前对话在前，归档在后。</summary>
        /// <returns>对话列表。</returns>
        public IReadOnlyList<AiConversation> GetAll()
        {
            lock (_syncRoot)
            {
                return _conversations
                    .OrderBy(item => item.Archived)
                    .ThenByDescending(item => item.UpdatedUtc, StringComparer.Ordinal)
                    .Select(Copy)
                    .ToList()
                    .AsReadOnly();
            }
        }

        /// <summary>按标识取一份副本。</summary>
        /// <param name="id">对话标识。</param>
        /// <returns>找不到时为空。</returns>
        public AiConversation Get(string id)
        {
            lock (_syncRoot)
            {
                return Copy(_conversations.FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)));
            }
        }

        /// <summary>新建一条还没有消息的对话。</summary>
        /// <param name="title">标题。为空时用「新对话」。</param>
        /// <param name="titleLocked">为 true 时，第一条问题不会改掉标题。</param>
        /// <returns>新建对话的副本。</returns>
        public AiConversation Create(string title, bool titleLocked)
        {
            AiConversation created = new AiConversation
            {
                Id = Guid.NewGuid().ToString("N"),
                Title = string.IsNullOrWhiteSpace(title) ? "新对话" : title.Trim(),
                TitleLocked = titleLocked,
                Archived = false,
                CreatedUtc = DateTime.UtcNow.ToString("o"),
                UpdatedUtc = DateTime.UtcNow.ToString("o"),
                Messages = new List<AiChatTurn>()
            };
            lock (_syncRoot)
            {
                _conversations.Add(Copy(created));
                WriteUnsafe(created);
            }

            Changed?.Invoke(this, EventArgs.Empty);
            return created;
        }

        /// <summary>修改标题并锁定，避免随后的问题再覆盖它。</summary>
        /// <param name="id">对话标识。</param>
        /// <param name="title">新标题。</param>
        public void Rename(string id, string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                throw new ArgumentException("对话标题不能为空。", nameof(title));
            Update(id, conversation =>
            {
                conversation.Title = title.Trim();
                conversation.TitleLocked = true;
            });
        }

        /// <summary>归档或取消归档。</summary>
        /// <param name="id">对话标识。</param>
        /// <param name="archived">归档时为 true。</param>
        public void SetArchived(string id, bool archived)
        {
            Update(id, conversation => conversation.Archived = archived);
        }

        /// <summary>追加一轮。标题仍是「新对话」且未锁定时，用这条用户消息的前几个字占位。</summary>
        /// <param name="id">对话标识。</param>
        /// <param name="turn">这一轮。用户轮次可带发送原文，助手轮次可带 token。</param>
        public void Append(string id, AiChatTurn turn)
        {
            if (turn == null) throw new ArgumentNullException(nameof(turn));
            Update(id, conversation =>
            {
                if (conversation.Messages == null)
                    conversation.Messages = new List<AiChatTurn>();
                conversation.Messages.Add(CopyTurn(turn));
                if (!conversation.TitleLocked
                    && string.Equals(conversation.Title, "新对话", StringComparison.Ordinal)
                    && string.Equals(turn.Role, "user", StringComparison.Ordinal))
                {
                    conversation.Title = ShortTitle(turn.Text);
                }
            });
        }

        /// <summary>用模型给出的标题替换现有标题。已锁定或整理后为空时不改。</summary>
        /// <param name="id">对话标识。</param>
        /// <param name="title">模型返回的原文。</param>
        /// <remarks>不设置 <see cref="AiConversation.TitleLocked"/>，后面的标题请求还可以再改。</remarks>
        public void SetGeneratedTitle(string id, string title)
        {
            string cleaned = NormalizeGeneratedTitle(title);
            if (cleaned == null) return;
            AiConversation current = Get(id);
            if (current == null || current.TitleLocked) return;
            if (string.Equals(current.Title, cleaned, StringComparison.Ordinal)) return;
            Update(id, conversation =>
            {
                if (conversation.TitleLocked) return;
                conversation.Title = cleaned;
            });
        }

        /// <summary>取模型返回的第一行，去掉引号，最长 18 个字。</summary>
        /// <param name="raw">模型原文。</param>
        /// <returns>可用标题。空的返回 null。</returns>
        public static string NormalizeGeneratedTitle(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            string line = raw.Trim();
            int breakAt = line.IndexOfAny(new[] { '\r', '\n' });
            if (breakAt >= 0)
                line = line.Substring(0, breakAt).Trim();
            line = line.Trim(' ', '\t', '"', '\'', '「', '」', '“', '”', '『', '』', '*', '#', '。', '.', '！', '!', '？', '?');
            if (line.Length == 0) return null;
            if (line.Length > 18)
                line = line.Substring(0, 18).Trim();
            return line.Length == 0 ? null : line;
        }

        /// <summary>删除对话文件。不删除审计日志。</summary>
        /// <param name="id">对话标识。</param>
        public void Delete(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            lock (_syncRoot)
            {
                _conversations.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                string path = FilePath(id);
                if (File.Exists(path))
                    File.Delete(path);
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        private void Update(string id, Action<AiConversation> change)
        {
            if (string.IsNullOrWhiteSpace(id))
                throw new ArgumentException("对话标识不能为空。", nameof(id));
            lock (_syncRoot)
            {
                AiConversation conversation = _conversations.FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                if (conversation == null)
                    throw new InvalidOperationException("对话不存在。");
                change(conversation);
                conversation.UpdatedUtc = DateTime.UtcNow.ToString("o");
                WriteUnsafe(conversation);
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        private List<AiConversation> Load()
        {
            List<AiConversation> conversations = new List<AiConversation>();
            if (!Directory.Exists(_paths.AiConversationsDirectory)) return conversations;
            foreach (string file in Directory.GetFiles(_paths.AiConversationsDirectory, "*.json"))
            {
                try
                {
                    AiConversation conversation = _serializer.Deserialize<AiConversation>(File.ReadAllText(file, Encoding.UTF8));
                    if (conversation == null) continue;
                    if (string.IsNullOrWhiteSpace(conversation.Id))
                        conversation.Id = Path.GetFileNameWithoutExtension(file);
                    if (string.IsNullOrWhiteSpace(conversation.Title))
                        conversation.Title = "新对话";
                    if (conversation.Messages == null)
                        conversation.Messages = new List<AiChatTurn>();
                    conversations.Add(Copy(conversation));
                }
                catch (Exception)
                {
                    AtomicFile.BackupCorrupt(file);
                }
            }

            return conversations;
        }

        private void WriteUnsafe(AiConversation conversation)
        {
            AtomicFile.Write(FilePath(conversation.Id), _serializer.Serialize(Copy(conversation)));
        }

        private string FilePath(string id)
        {
            return Path.Combine(_paths.AiConversationsDirectory, id + ".json");
        }

        private static string ShortTitle(string text)
        {
            string flat = (text ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
            if (flat.Length == 0) return "新对话";
            return flat.Length <= 18 ? flat : flat.Substring(0, 18);
        }

        private static AiConversation Copy(AiConversation source)
        {
            if (source == null) return null;
            AiConversation copy = new AiConversation
            {
                Id = source.Id,
                Title = source.Title,
                TitleLocked = source.TitleLocked,
                Archived = source.Archived,
                CreatedUtc = source.CreatedUtc,
                UpdatedUtc = source.UpdatedUtc,
                Messages = new List<AiChatTurn>()
            };
            if (source.Messages == null) return copy;
            foreach (AiChatTurn turn in source.Messages)
            {
                if (turn == null) continue;
                copy.Messages.Add(CopyTurn(turn));
            }

            return copy;
        }

        private static AiChatTurn CopyTurn(AiChatTurn turn)
        {
            AiChatTurn copy = new AiChatTurn
            {
                Role = turn.Role,
                Text = turn.Text,
                InputTokens = turn.InputTokens,
                OutputTokens = turn.OutputTokens,
                Sent = new List<AiSentSection>(),
                Call = CopyCall(turn.Call)
            };
            if (turn.Sent == null) return copy;
            foreach (AiSentSection section in turn.Sent)
            {
                if (section == null) continue;
                copy.Sent.Add(new AiSentSection
                {
                    Kind = section.Kind,
                    Title = section.Title,
                    Body = section.Body
                });
            }

            return copy;
        }

        /// <summary>复制调用信息，避免界面上的对象和文件里的对象互相改到。</summary>
        private static AiSentCall CopyCall(AiSentCall call)
        {
            if (call == null) return null;
            return new AiSentCall
            {
                Url = call.Url,
                Protocol = call.Protocol,
                ModelDisplayName = call.ModelDisplayName,
                ModelId = call.ModelId,
                SentUtc = call.SentUtc,
                HttpStatus = call.HttpStatus,
                ElapsedMs = call.ElapsedMs
            };
        }
    }
}
