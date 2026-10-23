using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>加载用户可编辑的思考强度到 API 参数映射。</summary>
    /// <remarks>
    /// 映射文件首次使用时从程序集内的默认 JSON 初始化；此后每次请求都重新读取，
    /// 以便用户编辑后立即生效。配置值只允许作为非内容请求参数合并，不得覆盖模型、
    /// 消息、系统说明、请求地址或认证信息。
    /// </remarks>
    public sealed class AiThinkingMappingStore
    {
        private const int CurrentVersion = 1;
        private const string DefaultResourceName = "DB2Sheet.Resources.ai-thinking-mappings.json";
        private static readonly object InitializationLock = new object();
        private static readonly HashSet<string> ReservedKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "model", "messages", "system", "contents", "systeminstruction", "prompt", "input",
            "authorization", "xapikey", "apikey", "key", "accesstoken", "token", "password",
            "secret", "secretkey", "credential"
        };

        private readonly ApplicationPaths _paths;

        /// <summary>创建映射存储，并在映射文件缺失时安装内嵌默认配置。</summary>
        /// <param name="paths">当前用户的本地数据路径。</param>
        public AiThinkingMappingStore(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            EnsureDefaultFile();
        }

        /// <summary>按协议、接口主机、型号和所选档位解析实际 API 参数。</summary>
        /// <param name="protocol">模型配置保存的协议。</param>
        /// <param name="endpoint">模型配置保存的接口地址。</param>
        /// <param name="modelId">本次选择的子型号。</param>
        /// <param name="level">automatic、low、medium 或 high。</param>
        /// <returns>可以合并到请求根对象的参数；automatic 始终为空。</returns>
        /// <exception cref="InvalidOperationException">配置损坏、无匹配规则或规则匹配不唯一。</exception>
        public IDictionary<string, object> Resolve(string protocol, string endpoint, string modelId, string level)
        {
            string normalizedLevel = NormalizeLevel(level);
            if (normalizedLevel == AiThinkingLevels.Automatic)
                return new Dictionary<string, object>(StringComparer.Ordinal);

            MappingDocument document = ReadDocument();
            ValidateDocument(document);

            string protocolName = AiProtocolNames.ToStorage(AiProtocolNames.Parse(protocol));
            Uri endpointUri;
            if (!Uri.TryCreate(endpoint ?? string.Empty, UriKind.Absolute, out endpointUri))
                throw new InvalidOperationException("模型接口地址无效，无法选择思考强度映射。");

            List<MappingRule> matches = document.Rules
                .Where(rule => string.Equals(rule.Protocol, protocolName, StringComparison.OrdinalIgnoreCase)
                    && WildcardMatches(rule.EndpointHost, endpointUri.Host)
                    && WildcardMatches(rule.ModelPattern, modelId ?? string.Empty))
                .ToList();
            if (matches.Count == 0)
                throw new InvalidOperationException(
                    "没有适用于当前协议、接口主机和型号的思考强度映射。请检查本地数据中的 ai-thinking-mappings.json。");

            int bestScore = matches.Max(Specificity);
            List<MappingRule> best = matches.Where(rule => Specificity(rule) == bestScore).ToList();
            if (best.Count != 1)
                throw new InvalidOperationException(
                    "ai-thinking-mappings.json 中有多条同等匹配的规则，请调整接口主机或型号匹配条件。");

            Dictionary<string, object> parameters;
            if (!best[0].Levels.TryGetValue(normalizedLevel, out parameters) || parameters == null)
                throw new InvalidOperationException(
                    "当前思考强度没有配置 API 参数，请检查 ai-thinking-mappings.json。");
            return new Dictionary<string, object>(parameters, StringComparer.Ordinal);
        }

        private void EnsureDefaultFile()
        {
            lock (InitializationLock)
            {
                if (File.Exists(_paths.AiThinkingMappingsFile))
                    return;

                Assembly assembly = typeof(AiThinkingMappingStore).Assembly;
                using (Stream stream = assembly.GetManifestResourceStream(DefaultResourceName))
                {
                    if (stream == null)
                        throw new InvalidOperationException("缺少内嵌思考强度映射默认配置。");
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                        AtomicFile.Write(_paths.AiThinkingMappingsFile, reader.ReadToEnd());
                }
            }
        }

        private MappingDocument ReadDocument()
        {
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer { MaxJsonLength = 1024 * 1024 };
                string json = File.ReadAllText(_paths.AiThinkingMappingsFile, Encoding.UTF8);
                MappingDocument document = serializer.Deserialize<MappingDocument>(json);
                if (document == null)
                    throw new InvalidOperationException("配置文件内容为空。");
                return document;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    "读取思考强度映射配置失败：" + exception.Message,
                    exception);
            }
        }

        private static void ValidateDocument(MappingDocument document)
        {
            if (document.Version != CurrentVersion)
                throw new InvalidOperationException(
                    "ai-thinking-mappings.json 的 version 必须为 " + CurrentVersion.ToString() + "。");
            if (document.Rules == null || document.Rules.Count == 0)
                throw new InvalidOperationException("ai-thinking-mappings.json 至少需要一条规则。");

            foreach (MappingRule rule in document.Rules)
            {
                if (rule == null)
                    throw new InvalidOperationException("ai-thinking-mappings.json 不能包含空规则。");
                string protocol = AiProtocolNames.ToStorage(AiProtocolNames.Parse(rule.Protocol));
                if (!string.Equals(protocol, rule.Protocol, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("思考强度映射中的 protocol 必须是 openai、claude 或 gemini。");
                if (string.IsNullOrWhiteSpace(rule.EndpointHost))
                    throw new InvalidOperationException("思考强度映射的 endpointHost 不能为空。");
                if (string.IsNullOrWhiteSpace(rule.ModelPattern))
                    rule.ModelPattern = "*";
                if (rule.Levels == null)
                    throw new InvalidOperationException("思考强度映射必须包含 levels。");

                foreach (string level in new[] { AiThinkingLevels.Low, AiThinkingLevels.Medium, AiThinkingLevels.High })
                {
                    Dictionary<string, object> parameters;
                    if (!rule.Levels.TryGetValue(level, out parameters) || parameters == null)
                        throw new InvalidOperationException("思考强度映射缺少档位：" + level + "。");
                    ValidateParameters(parameters);
                }

                if (rule.Levels.ContainsKey(AiThinkingLevels.Automatic))
                    throw new InvalidOperationException("automatic 档位固定不发送思考参数，请从 levels 中移除该项。");
            }
        }

        private static void ValidateParameters(IDictionary<string, object> parameters)
        {
            foreach (KeyValuePair<string, object> parameter in parameters)
            {
                if (string.IsNullOrWhiteSpace(parameter.Key) || ReservedKeys.Contains(NormalizeKey(parameter.Key)))
                    throw new InvalidOperationException(
                        "思考强度映射包含受保护或无效的请求参数：" + (parameter.Key ?? string.Empty) + "。");
                ValidateJsonValue(parameter.Value);
            }
        }

        private static void ValidateJsonValue(object value)
        {
            if (value == null || value is string || value is bool || value is int || value is long
                || value is double || value is decimal)
                return;

            IDictionary<string, object> dictionary = value as IDictionary<string, object>;
            if (dictionary != null)
            {
                ValidateParameters(dictionary);
                return;
            }

            IEnumerable sequence = value as IEnumerable;
            if (sequence != null)
            {
                foreach (object item in sequence)
                    ValidateJsonValue(item);
                return;
            }

            throw new InvalidOperationException("思考强度映射只允许 JSON 标量、对象和数组。");
        }

        private static int Specificity(MappingRule rule)
        {
            return LiteralLength(rule.EndpointHost) + LiteralLength(rule.ModelPattern);
        }

        private static int LiteralLength(string pattern)
        {
            return (pattern ?? string.Empty).Count(character => character != '*' && character != '?');
        }

        private static bool WildcardMatches(string pattern, string value)
        {
            string candidate = pattern ?? "*";
            int patternIndex = 0;
            int valueIndex = 0;
            int starIndex = -1;
            int retryValueIndex = -1;
            while (valueIndex < (value ?? string.Empty).Length)
            {
                string text = value ?? string.Empty;
                if (patternIndex < candidate.Length
                    && (candidate[patternIndex] == '?'
                        || char.ToUpperInvariant(candidate[patternIndex]) == char.ToUpperInvariant(text[valueIndex])))
                {
                    patternIndex++;
                    valueIndex++;
                }
                else if (patternIndex < candidate.Length && candidate[patternIndex] == '*')
                {
                    starIndex = patternIndex++;
                    retryValueIndex = valueIndex;
                }
                else if (starIndex >= 0)
                {
                    patternIndex = starIndex + 1;
                    valueIndex = ++retryValueIndex;
                }
                else
                {
                    return false;
                }
            }

            while (patternIndex < candidate.Length && candidate[patternIndex] == '*')
                patternIndex++;
            return patternIndex == candidate.Length;
        }

        private static string NormalizeLevel(string level)
        {
            if (string.Equals(level, AiThinkingLevels.Automatic, StringComparison.OrdinalIgnoreCase)
                || string.Equals(level, AiThinkingLevels.Low, StringComparison.OrdinalIgnoreCase)
                || string.Equals(level, AiThinkingLevels.Medium, StringComparison.OrdinalIgnoreCase)
                || string.Equals(level, AiThinkingLevels.High, StringComparison.OrdinalIgnoreCase))
                return level.ToLowerInvariant();
            throw new ArgumentException("未知的思考强度档位。", nameof(level));
        }

        private static string NormalizeKey(string key)
        {
            return (key ?? string.Empty)
                .Replace("_", string.Empty)
                .Replace("-", string.Empty)
                .Replace(".", string.Empty)
                .ToLowerInvariant();
        }

        /// <summary>映射文件的根对象。</summary>
        public sealed class MappingDocument
        {
            /// <summary>获取映射配置的格式版本。</summary>
            public int Version { get; set; }
            /// <summary>获取协议、接口主机和型号规则。</summary>
            public List<MappingRule> Rules { get; set; }
        }

        /// <summary>一条按协议、主机和子型号匹配的 API 参数规则。</summary>
        public sealed class MappingRule
        {
            /// <summary>获取请求协议。</summary>
            public string Protocol { get; set; }
            /// <summary>获取接口主机通配模式。</summary>
            public string EndpointHost { get; set; }
            /// <summary>获取子型号通配模式。</summary>
            public string ModelPattern { get; set; }
            /// <summary>获取每个手动档位对应的 API 参数。</summary>
            public Dictionary<string, Dictionary<string, object>> Levels { get; set; }
        }
    }

}
