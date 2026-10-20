using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Net.Http;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>检查并描述用户填写的模型地址。查询参数值不进入说明。</summary>
    public static class AiEndpoints
    {
        private static readonly HashSet<string> SecretQueryNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "key", "api_key", "apikey", "api-key", "access_token", "token", "secret", "password", "pwd"
        };

        /// <summary>拒绝不是 http(s)、或把密钥写在地址里的输入。</summary>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <exception cref="ArgumentException">地址不能用。</exception>
        public static void EnsureUsable(string endpoint)
        {
            if (!Uri.TryCreate((endpoint ?? string.Empty).Trim(), UriKind.Absolute, out Uri uri))
                throw new ArgumentException("接口地址必须是完整的 http 或 https 地址。");
            if (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
                throw new ArgumentException("接口地址只允许 http 或 https。");
            if (!string.IsNullOrEmpty(uri.UserInfo))
                throw new ArgumentException("请把密钥填在密钥框里，不要写进地址。");
        }

        /// <summary>给审计日志用的主机、路径和查询参数名。</summary>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <returns>不含查询参数值的说明。地址无法解析时返回脱敏后的原文。</returns>
        public static string Describe(string endpoint)
        {
            if (!Uri.TryCreate((endpoint ?? string.Empty).Trim(), UriKind.Absolute, out Uri uri))
                return AiTextRedactor.Redact(endpoint, out string _);
            string names = string.Join(",", QueryNames(uri));
            return uri.Host + uri.AbsolutePath + (names.Length == 0 ? string.Empty : " query=" + names);
        }

        /// <summary>去掉查询串里的密钥参数，保留 api-version 这类普通参数。</summary>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <returns>可以实际请求的地址。</returns>
        public static string StripSecretQuery(string endpoint)
        {
            Uri uri = new Uri(endpoint.Trim());
            if (string.IsNullOrEmpty(uri.Query))
                return uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
            List<string> kept = new List<string>();
            string query = uri.Query.TrimStart('?');
            foreach (string pair in query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string name = pair;
                int separator = pair.IndexOf('=');
                if (separator >= 0)
                    name = pair.Substring(0, separator);
                if (SecretQueryNames.Contains(Uri.UnescapeDataString(name)))
                    continue;
                kept.Add(pair);
            }

            string path = uri.GetLeftPart(UriPartial.Path).TrimEnd('/');
            return kept.Count == 0 ? path : path + "?" + string.Join("&", kept);
        }

        private static IEnumerable<string> QueryNames(Uri uri)
        {
            if (string.IsNullOrEmpty(uri.Query)) yield break;
            string query = uri.Query.TrimStart('?');
            foreach (string pair in query.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int separator = pair.IndexOf('=');
                string name = separator >= 0 ? pair.Substring(0, separator) : pair;
                yield return Uri.UnescapeDataString(name);
            }
        }
    }

    /// <summary>按三种协议把已经拼好的消息发给用户指定的地址。</summary>
    /// <remarks>不知道数据库连接。调用方负责脱敏，并在发送前写好审计日志。</remarks>
    public sealed class AiChatClient
    {
        private static readonly HttpClient Http = new HttpClient { Timeout = TimeSpan.FromSeconds(120) };
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        /// <summary>一次尚未发出的 HTTP 调用。标头里的密钥不会写入审计日志。</summary>
        public sealed class OutboundCall
        {
            /// <summary>实际请求地址。</summary>
            public string Url { get; set; }

            /// <summary>请求头。值可能含密钥，只用于发送。</summary>
            public List<KeyValuePair<string, string>> Headers { get; set; }

            /// <summary>请求体 JSON。</summary>
            public string Body { get; set; }
        }

        /// <summary>模型返回的 HTTP 结果。正文尚未解析。</summary>
        public sealed class HttpResult
        {
            /// <summary>创建结果。</summary>
            /// <param name="statusCode">HTTP 状态码。</param>
            /// <param name="body">响应正文。</param>
            public HttpResult(int statusCode, string body)
            {
                StatusCode = statusCode;
                Body = body ?? string.Empty;
            }

            /// <summary>HTTP 状态码。</summary>
            public int StatusCode { get; }

            /// <summary>响应正文。</summary>
            public string Body { get; }
        }

        /// <summary>按协议拼出请求。不访问网络。</summary>
        /// <param name="protocol">三种协议之一。</param>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <param name="apiKey">本次请求的明文密钥。</param>
        /// <param name="modelName">模型名。</param>
        /// <param name="systemText">系统说明。测试请求可以为空。</param>
        /// <param name="turns">user 与 assistant 往来，最后一轮是本次用户消息。</param>
        /// <returns>待发送的调用。</returns>
        public OutboundCall Build(
            AiProtocol protocol,
            string endpoint,
            string apiKey,
            string modelName,
            string systemText,
            IList<AiChatTurn> turns)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("请先填写密钥。");
            if (string.IsNullOrWhiteSpace(modelName))
                throw new InvalidOperationException("请先填写模型名。");
            string model = modelName.Trim();
            string url = ChatUrl(protocol, endpoint, model);
            switch (protocol)
            {
                case AiProtocol.Claude:
                    return BuildClaude(url, apiKey, model, systemText, turns);
                case AiProtocol.Gemini:
                    return BuildGemini(url, apiKey, model, systemText, turns);
                default:
                    return BuildOpenAi(url, apiKey, model, systemText, turns);
            }
        }

        /// <summary>拼出这次实际要 POST 的地址。发送和请求详情用这一份，避免两套路径。</summary>
        /// <param name="protocol">三种协议之一。</param>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <param name="modelName">Gemini 把子型号写进路径。其他协议不使用。</param>
        /// <returns>已经去掉密钥查询参数的地址。</returns>
        /// <exception cref="ArgumentException">地址不能用。</exception>
        public static string ChatUrl(AiProtocol protocol, string endpoint, string modelName)
        {
            AiEndpoints.EnsureUsable(endpoint);
            string clean = AiEndpoints.StripSecretQuery(endpoint);
            switch (protocol)
            {
                case AiProtocol.Claude:
                    return ClaudeUrl(clean);
                case AiProtocol.Gemini:
                    return GeminiGenerateUrl(clean, (modelName ?? string.Empty).Trim());
                default:
                    return OpenAiUrl(clean);
            }
        }

        /// <summary>向厂商拉取可调用的型号。不写审计，调用方在探测对话前再写。</summary>
        /// <param name="protocol">三种协议之一。</param>
        /// <param name="endpoint">用户填写的地址。</param>
        /// <param name="apiKey">本次请求的明文密钥。</param>
        /// <returns>去重后的型号。空列表不会返回，改为抛出。</returns>
        /// <remarks>等待网络时让出调用线程。地址由聊天地址推出来，密钥只放在请求头里。</remarks>
        public async System.Threading.Tasks.Task<IReadOnlyList<string>> ListModelsAsync(
            AiProtocol protocol,
            string endpoint,
            string apiKey)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new InvalidOperationException("请先填写密钥。");
            AiEndpoints.EnsureUsable(endpoint);
            string url = ModelsUrl(protocol, AiEndpoints.StripSecretQuery(endpoint));
            List<KeyValuePair<string, string>> headers = AuthHeaders(protocol, apiKey);
            HttpResult result = await GetAsync(url, headers).ConfigureAwait(false);
            if (result.StatusCode < 200 || result.StatusCode >= 300)
                throw new InvalidOperationException(
                    "读取型号失败（HTTP " + result.StatusCode.ToString() + "）。" + ReadError(result.Body));
            IReadOnlyList<string> ids = ReadModelIds(result.Body);
            if (ids.Count == 0)
                throw new InvalidOperationException("接口没有返回型号。");
            return ids;
        }

        /// <summary>从聊天地址推出型号列表地址。</summary>
        /// <param name="protocol">协议。</param>
        /// <param name="endpoint">已经去掉密钥查询参数的地址。</param>
        /// <returns>可直接 GET 的地址。</returns>
        public static string ModelsUrl(AiProtocol protocol, string endpoint)
        {
            Uri uri = new Uri(endpoint);
            string path = uri.AbsolutePath ?? string.Empty;
            switch (protocol)
            {
                case AiProtocol.Claude:
                    path = SwapSuffix(path, "/messages", "/models");
                    if (!EndsWith(path, "/models"))
                        path = EndsWith(path, "/v1") ? path + "/models" : TrimSlash(path) + "/v1/models";
                    break;
                case AiProtocol.Gemini:
                    int models = IndexOf(path, "/models");
                    if (models >= 0)
                        path = path.Substring(0, models + "/models".Length);
                    else if (EndsWith(path, "/v1beta") || EndsWith(path, "/v1"))
                        path = path + "/models";
                    else
                        path = TrimSlash(path) + "/v1beta/models";
                    break;
                default:
                    path = SwapSuffix(path, "/chat/completions", "/models");
                    if (!EndsWith(path, "/models"))
                        path = TrimSlash(path) + "/models";
                    break;
            }

            UriBuilder builder = new UriBuilder(uri) { Path = path };
            return builder.Uri.AbsoluteUri;
        }

        /// <summary>发送已经拼好的调用。</summary>
        /// <param name="call">由 <see cref="Build"/> 产生的调用。</param>
        /// <returns>状态码和响应正文。</returns>
        /// <remarks>在等待响应时让出调用线程。不要在后台线程接着访问 Excel。</remarks>
        public async System.Threading.Tasks.Task<HttpResult> SendAsync(OutboundCall call)
        {
            if (call == null) throw new ArgumentNullException(nameof(call));
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, call.Url))
            {
                request.Content = new StringContent(call.Body ?? string.Empty, Encoding.UTF8, "application/json");
                if (call.Headers != null)
                {
                    foreach (KeyValuePair<string, string> header in call.Headers)
                        request.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }

                using (HttpResponseMessage response = await Http.SendAsync(request).ConfigureAwait(false))
                {
                    string body = response.Content == null
                        ? string.Empty
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return new HttpResult((int)response.StatusCode, body);
                }
            }
        }

        /// <summary>从成功响应里取出助手文字。</summary>
        /// <param name="protocol">与发送时相同的协议。</param>
        /// <param name="body">响应 JSON。</param>
        /// <returns>助手文字。</returns>
        /// <exception cref="InvalidOperationException">响应里没有可显示的文本。</exception>
        public string ReadReply(AiProtocol protocol, string body)
        {
            object parsed = _serializer.DeserializeObject(body ?? string.Empty);
            Dictionary<string, object> root = parsed as Dictionary<string, object>;
            if (root == null)
                throw new InvalidOperationException("模型没有返回可识别的内容。");

            string reply;
            switch (protocol)
            {
                case AiProtocol.Claude:
                    reply = ReadClaude(root);
                    break;
                case AiProtocol.Gemini:
                    reply = ReadGemini(root);
                    break;
                default:
                    reply = ReadOpenAi(root);
                    break;
            }

            if (string.IsNullOrWhiteSpace(reply))
                throw new InvalidOperationException("模型没有返回文本。");
            return reply;
        }

        /// <summary>从成功响应里取出输入和输出 token。缺字段或类型不对时对应一侧为空。</summary>
        /// <param name="protocol">与发送时相同的协议。</param>
        /// <param name="body">响应 JSON。</param>
        /// <returns>两侧都可能为空。解析失败时两侧都为空，不抛异常。</returns>
        /// <remarks>用量不是回复正文。解析失败时这一轮不显示计数。</remarks>
        public AiReply ReadUsage(AiProtocol protocol, string body)
        {
            AiReply usage = new AiReply();
            try
            {
                Dictionary<string, object> root = _serializer.DeserializeObject(body ?? string.Empty) as Dictionary<string, object>;
                if (root == null) return usage;
                Dictionary<string, object> bag;
                string inputName;
                string outputName;
                switch (protocol)
                {
                    case AiProtocol.Claude:
                        bag = Dict(root, "usage");
                        inputName = "input_tokens";
                        outputName = "output_tokens";
                        break;
                    case AiProtocol.Gemini:
                        bag = Dict(root, "usageMetadata");
                        inputName = "promptTokenCount";
                        outputName = "candidatesTokenCount";
                        break;
                    default:
                        bag = Dict(root, "usage");
                        inputName = "prompt_tokens";
                        outputName = "completion_tokens";
                        break;
                }

                usage.InputTokens = Count(bag, inputName);
                usage.OutputTokens = Count(bag, outputName);
            }
            catch (Exception)
            {
                // 用量不是回复正文。解析失败时这一轮不显示计数。
            }

            return usage;
        }

        /// <summary>从错误响应里取一段可读说明。</summary>
        /// <param name="body">响应正文。</param>
        /// <returns>错误说明。解析不到时返回截短后的正文。</returns>
        public string ReadError(string body)
        {
            try
            {
                Dictionary<string, object> root = _serializer.DeserializeObject(body ?? string.Empty) as Dictionary<string, object>;
                Dictionary<string, object> error = root == null ? null : root.ContainsKey("error") ? root["error"] as Dictionary<string, object> : null;
                if (error != null && error.ContainsKey("message") && error["message"] != null)
                    return Convert.ToString(error["message"]);
            }
            catch (Exception)
            {
                // 错误正文不一定是 JSON。下面返回截短原文。
            }

            string text = body ?? string.Empty;
            return text.Length <= 2000 ? text : text.Substring(0, 2000);
        }

        private OutboundCall BuildOpenAi(string url, string apiKey, string modelName, string systemText, IList<AiChatTurn> turns)
        {
            List<Dictionary<string, string>> messages = new List<Dictionary<string, string>>();
            if (!string.IsNullOrWhiteSpace(systemText))
                messages.Add(Message("system", systemText));
            foreach (AiChatTurn turn in Merge(turns))
                messages.Add(Message(turn.Role, turn.Text));
            Dictionary<string, object> payload = new Dictionary<string, object>
            {
                { "model", modelName },
                { "messages", messages }
            };
            return new OutboundCall
            {
                Url = url,
                Body = _serializer.Serialize(payload),
                Headers = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("Authorization", "Bearer " + apiKey)
                }
            };
        }

        private OutboundCall BuildClaude(string url, string apiKey, string modelName, string systemText, IList<AiChatTurn> turns)
        {
            List<Dictionary<string, string>> messages = new List<Dictionary<string, string>>();
            foreach (AiChatTurn turn in Merge(turns))
            {
                string role = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "assistant" : "user";
                messages.Add(Message(role, turn.Text));
            }

            Dictionary<string, object> payload = new Dictionary<string, object>
            {
                { "model", modelName },
                { "max_tokens", 4096 },
                { "messages", messages }
            };
            if (!string.IsNullOrWhiteSpace(systemText))
                payload["system"] = systemText;
            return new OutboundCall
            {
                Url = url,
                Body = _serializer.Serialize(payload),
                Headers = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("x-api-key", apiKey),
                    new KeyValuePair<string, string>("anthropic-version", "2023-06-01")
                }
            };
        }

        private OutboundCall BuildGemini(string url, string apiKey, string modelName, string systemText, IList<AiChatTurn> turns)
        {
            List<Dictionary<string, object>> contents = new List<Dictionary<string, object>>();
            foreach (AiChatTurn turn in Merge(turns))
            {
                string role = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "model" : "user";
                contents.Add(new Dictionary<string, object>
                {
                    { "role", role },
                    { "parts", new object[] { new Dictionary<string, string> { { "text", turn.Text ?? string.Empty } } } }
                });
            }

            Dictionary<string, object> payload = new Dictionary<string, object> { { "contents", contents } };
            if (!string.IsNullOrWhiteSpace(systemText))
            {
                payload["systemInstruction"] = new Dictionary<string, object>
                {
                    { "parts", new object[] { new Dictionary<string, string> { { "text", systemText } } } }
                };
            }

            return new OutboundCall
            {
                Url = url,
                Body = _serializer.Serialize(payload),
                Headers = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("x-goog-api-key", apiKey)
                }
            };
        }

        private static string ReadOpenAi(Dictionary<string, object> root)
        {
            Dictionary<string, object> message = First(root, "choices") as Dictionary<string, object>;
            if (message != null && message.ContainsKey("message"))
                message = message["message"] as Dictionary<string, object>;
            return message == null ? null : TextOf(message.ContainsKey("content") ? message["content"] : null);
        }

        private static string ReadClaude(Dictionary<string, object> root)
        {
            return TextOf(root.ContainsKey("content") ? root["content"] : null);
        }

        private static string ReadGemini(Dictionary<string, object> root)
        {
            Dictionary<string, object> candidate = First(root, "candidates") as Dictionary<string, object>;
            if (candidate == null || !candidate.ContainsKey("content")) return null;
            Dictionary<string, object> content = candidate["content"] as Dictionary<string, object>;
            if (content == null || !content.ContainsKey("parts")) return null;
            return TextOf(content["parts"]);
        }

        private static Dictionary<string, object> Dict(Dictionary<string, object> root, string name)
        {
            if (root == null || !root.ContainsKey(name)) return null;
            return root[name] as Dictionary<string, object>;
        }

        /// <summary>读一个非负整数计数。小数、负数或超出 Int32 时当作没有。</summary>
        private static int? Count(Dictionary<string, object> bag, string name)
        {
            if (bag == null || !bag.ContainsKey(name) || bag[name] == null) return null;
            try
            {
                decimal number = Convert.ToDecimal(bag[name], CultureInfo.InvariantCulture);
                if (number < 0 || number > int.MaxValue || number != decimal.Truncate(number))
                    return null;
                return (int)number;
            }
            catch (Exception)
            {
                // 计数不是整数时当作没有这一侧，不让用量解析弄失败回复。
                return null;
            }
        }

        private static object First(Dictionary<string, object> root, string name)
        {
            if (root == null || !root.ContainsKey(name)) return null;
            IList list = root[name] as IList;
            if (list == null || list.Count == 0) return null;
            return list[0];
        }

        private static string TextOf(object value)
        {
            if (value == null) return null;
            string text = value as string;
            if (text != null) return text;
            IList list = value as IList;
            if (list == null) return null;
            StringBuilder builder = new StringBuilder();
            foreach (object item in list)
            {
                string direct = item as string;
                if (direct != null)
                {
                    builder.Append(direct);
                    continue;
                }

                Dictionary<string, object> part = item as Dictionary<string, object>;
                if (part != null && part.ContainsKey("text") && part["text"] != null)
                    builder.Append(Convert.ToString(part["text"]));
            }

            return builder.ToString();
        }

        private async System.Threading.Tasks.Task<HttpResult> GetAsync(string url, List<KeyValuePair<string, string>> headers)
        {
            using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, url))
            {
                if (headers != null)
                {
                    foreach (KeyValuePair<string, string> header in headers)
                        request.Headers.TryAddWithoutValidation(header.Key, header.Value);
                }

                using (HttpResponseMessage response = await Http.SendAsync(request).ConfigureAwait(false))
                {
                    string body = response.Content == null
                        ? string.Empty
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return new HttpResult((int)response.StatusCode, body);
                }
            }
        }

        private static List<KeyValuePair<string, string>> AuthHeaders(AiProtocol protocol, string apiKey)
        {
            switch (protocol)
            {
                case AiProtocol.Claude:
                    return new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("x-api-key", apiKey),
                        new KeyValuePair<string, string>("anthropic-version", "2023-06-01")
                    };
                case AiProtocol.Gemini:
                    return new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("x-goog-api-key", apiKey)
                    };
                default:
                    return new List<KeyValuePair<string, string>>
                    {
                        new KeyValuePair<string, string>("Authorization", "Bearer " + apiKey)
                    };
            }
        }

        private IReadOnlyList<string> ReadModelIds(string body)
        {
            object parsed = _serializer.DeserializeObject(body ?? string.Empty);
            Dictionary<string, object> root = parsed as Dictionary<string, object>;
            if (root == null)
                throw new InvalidOperationException("模型列表不是可识别的内容。");
            List<string> ids = new List<string>();
            CollectIds(ids, root, "data");
            CollectIds(ids, root, "models");
            return ids;
        }

        private static void CollectIds(List<string> ids, Dictionary<string, object> root, string key)
        {
            if (root == null || !root.ContainsKey(key)) return;
            IList list = root[key] as IList;
            if (list == null) return;
            foreach (object item in list)
            {
                string direct = item as string;
                if (!string.IsNullOrWhiteSpace(direct))
                {
                    AddModelId(ids, direct);
                    continue;
                }

                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null) continue;
                string id = TextOf(row.ContainsKey("id") ? row["id"] : null);
                if (string.IsNullOrWhiteSpace(id))
                    id = TextOf(row.ContainsKey("name") ? row["name"] : null);
                AddModelId(ids, id);
            }
        }

        private static void AddModelId(List<string> ids, string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            string trimmed = id.Trim();
            const string prefix = "models/";
            if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                trimmed = trimmed.Substring(prefix.Length);
            if (trimmed.Length == 0) return;
            foreach (string existing in ids)
            {
                if (string.Equals(existing, trimmed, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            ids.Add(trimmed);
        }

        /// <summary>OpenAI 兼容的 chat/completions 地址。已经带这个后缀时不再追加。</summary>
        private static string OpenAiUrl(string endpoint)
        {
            return endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase)
                ? endpoint
                : endpoint + "/chat/completions";
        }

        /// <summary>Claude Messages 地址。已是 /messages 时保持不变，以 /v1 结尾时只补 /messages。</summary>
        private static string ClaudeUrl(string endpoint)
        {
            return endpoint.EndsWith("/messages", StringComparison.OrdinalIgnoreCase)
                ? endpoint
                : endpoint.EndsWith("/v1", StringComparison.OrdinalIgnoreCase)
                    ? endpoint + "/messages"
                    : endpoint + "/v1/messages";
        }

        /// <summary>用所选子型号拼 Gemini 的 generateContent 地址，不沿用地址里写死的旧型号。</summary>
        private static string GeminiGenerateUrl(string endpoint, string modelName)
        {
            Uri uri = new Uri(endpoint);
            string path = uri.AbsolutePath ?? string.Empty;
            int models = IndexOf(path, "/models");
            string prefix = models >= 0 ? path.Substring(0, models) : TrimSlash(path);
            if (!EndsWith(prefix, "/v1beta") && !EndsWith(prefix, "/v1"))
                prefix = TrimSlash(prefix) + "/v1beta";
            string next = prefix + "/models/" + Uri.EscapeDataString(modelName) + ":generateContent";
            UriBuilder builder = new UriBuilder(uri) { Path = next };
            return builder.Uri.AbsoluteUri;
        }

        private static string SwapSuffix(string path, string suffix, string replacement)
        {
            if (!EndsWith(path, suffix)) return path;
            return path.Substring(0, path.Length - suffix.Length) + replacement;
        }

        private static bool EndsWith(string path, string suffix)
        {
            return path != null && path.EndsWith(suffix, StringComparison.OrdinalIgnoreCase);
        }

        private static int IndexOf(string path, string value)
        {
            return path == null ? -1 : path.IndexOf(value, StringComparison.OrdinalIgnoreCase);
        }

        private static string TrimSlash(string path)
        {
            return (path ?? string.Empty).TrimEnd('/');
        }

        private static Dictionary<string, string> Message(string role, string text)
        {
            return new Dictionary<string, string>
            {
                { "role", role },
                { "content", text ?? string.Empty }
            };
        }

        private static List<AiChatTurn> Merge(IList<AiChatTurn> turns)
        {
            List<AiChatTurn> merged = new List<AiChatTurn>();
            if (turns == null) return merged;
            foreach (AiChatTurn turn in turns)
            {
                if (turn == null || string.IsNullOrWhiteSpace(turn.Text)) continue;
                string role = string.Equals(turn.Role, "assistant", StringComparison.OrdinalIgnoreCase) ? "assistant" : "user";
                if (merged.Count > 0 && string.Equals(merged[merged.Count - 1].Role, role, StringComparison.Ordinal))
                {
                    merged[merged.Count - 1].Text = merged[merged.Count - 1].Text + "\r\n" + turn.Text;
                    continue;
                }

                merged.Add(new AiChatTurn { Role = role, Text = turn.Text });
            }

            return merged;
        }
    }
}
