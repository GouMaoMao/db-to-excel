using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Providers
{
    /// <summary>启动 JDBC 转接进程时使用的 Java 路径。</summary>
    internal sealed class JdbcBridgeLaunch
    {
        /// <summary>创建启动参数。</summary>
        /// <param name="javaExecutable">用户指定的 java.exe；为空时自动搜索本机。</param>
        public JdbcBridgeLaunch(string javaExecutable)
        {
            JavaExecutable = javaExecutable ?? string.Empty;
        }

        /// <summary>获取用户指定的 Java 路径。</summary>
        public string JavaExecutable { get; }
    }

    /// <summary>从插件程序集中释放 JDBC 转接程序，供 java.exe 启动。</summary>
    /// <remarks>
    /// jar 嵌在 DLL 内，避免 Excel 从 VSTO 缓存加载时旁边没有这个文件。
    /// 释放位置由调用方指定，通常是本地数据目录，不是连接方案里的厂商 jar。
    /// </remarks>
    internal static class JdbcBridgeFiles
    {
        /// <summary>嵌入资源名。</summary>
        public const string ResourceName = "DB2Sheet.db2sheet-jdbc-bridge.jar";

        /// <summary>释放后的文件名。</summary>
        public const string FileName = "db2sheet-jdbc-bridge.jar";

        /// <summary>程序集里没有转接程序时的说明。这不是连接方案中的厂商 jar。</summary>
        public const string MissingBridgeMessage = "插件内缺少 JDBC 转接程序。请重新编译或重新安装插件。这不是连接方案中的驱动 jar。";

        private static readonly object ExtractGate = new object();

        /// <summary>把嵌入的转接程序释放到指定目录。内容一致时不重写。</summary>
        /// <param name="directory">目标目录，会在不存在时创建。</param>
        /// <returns>可供 java -jar 使用的 jar 路径。</returns>
        /// <exception cref="InvalidOperationException">嵌入资源缺失或无法写入目标文件。</exception>
        public static string EnsureExtracted(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentException("转接程序目录不能为空。", nameof(directory));
            }

            byte[] embedded = ReadEmbedded();
            if (embedded == null || embedded.Length == 0)
            {
                throw new InvalidOperationException(MissingBridgeMessage);
            }

            lock (ExtractGate)
            {
                try
                {
                    Directory.CreateDirectory(directory);
                    string path = Path.Combine(directory, FileName);
                    if (File.Exists(path) && SameContent(path, embedded))
                    {
                        return path;
                    }

                    string temporary = path + ".tmp";
                    File.WriteAllBytes(temporary, embedded);
                    try
                    {
                        if (File.Exists(path))
                        {
                            File.Replace(temporary, path, null);
                        }
                        else
                        {
                            File.Move(temporary, path);
                        }
                    }
                    finally
                    {
                        if (File.Exists(temporary))
                        {
                            File.Delete(temporary);
                        }
                    }

                    return path;
                }
                catch (InvalidOperationException)
                {
                    throw;
                }
                catch (Exception exception)
                {
                    throw new InvalidOperationException("无法释放 JDBC 转接程序：" + exception.Message, exception);
                }
            }
        }

        private static byte[] ReadEmbedded()
        {
            using (Stream stream = typeof(JdbcBridgeFiles).Assembly.GetManifestResourceStream(ResourceName))
            {
                if (stream == null) return null;
                using (MemoryStream memory = new MemoryStream())
                {
                    stream.CopyTo(memory);
                    return memory.ToArray();
                }
            }
        }

        private static bool SameContent(string path, byte[] embedded)
        {
            FileInfo info = new FileInfo(path);
            if (info.Length != embedded.Length) return false;
            byte[] existing = File.ReadAllBytes(path);
            if (existing.Length != embedded.Length) return false;
            for (int index = 0; index < existing.Length; index++)
            {
                if (existing[index] != embedded[index]) return false;
            }

            return true;
        }
    }

    /// <summary>维护一个长期运行的 Java 转接进程，并按请求标识分发响应。</summary>
    /// <remarks>
    /// 标准输入可能包含密码，因此本类不记录请求原文。进程在加载项退出或 Java 路径变化且没有进行中的查询时重启。
    /// 厂商驱动 jar 随每条请求传入，不作为进程启动参数。调用方负责在不再使用时调用 <see cref="Dispose"/> 结束进程。
    /// </remarks>
    internal sealed class JdbcBridgeHost : IDisposable
    {
        private readonly ILogger _logger;
        private readonly string _bridgeDirectory;
        private readonly object _gate = new object();
        private readonly object _writeGate = new object();
        private readonly SemaphoreSlim _lifecycle = new SemaphoreSlim(1, 1);
        private readonly Dictionary<string, JdbcReplyQueue> _queues =
            new Dictionary<string, JdbcReplyQueue>(StringComparer.Ordinal);
        private Process _process;
        private StreamWriter _input;
        private string _signature = string.Empty;
        private int _generation;
        private int _activeQueries;
        private bool _disposed;

        /// <summary>创建转接进程宿主。</summary>
        /// <param name="logger">诊断日志；为空时不写日志。</param>
        /// <param name="bridgeDirectory">释放嵌入转接程序的目录，通常是本地数据根目录。</param>
        public JdbcBridgeHost(ILogger logger, string bridgeDirectory)
        {
            _logger = logger;
            _bridgeDirectory = bridgeDirectory ?? throw new ArgumentNullException(nameof(bridgeDirectory));
        }

        /// <summary>发送一次会返回单帧结果的操作，例如测试连接或读取目录。</summary>
        /// <param name="launch">Java 启动参数。</param>
        /// <param name="request">请求内容。密码只写入标准输入。</param>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>终态响应。</returns>
        /// <exception cref="InvalidOperationException">Java、转接程序或驱动操作失败。</exception>
        /// <exception cref="OperationCanceledException">等待被取消。</exception>
        public async Task<JdbcBridgeReply> RequestAsync(
            JdbcBridgeLaunch launch,
            JdbcBridgeRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null) throw new ArgumentNullException(nameof(request));
            await EnsureReadyAsync(launch, cancellationToken).ConfigureAwait(false);
            if (string.IsNullOrEmpty(request.Id))
            {
                request.Id = Guid.NewGuid().ToString("N");
            }

            JdbcReplyQueue queue = Register(request.Id);
            try
            {
                Write(request);
                while (true)
                {
                    JdbcBridgeReply reply = await queue.DequeueAsync(cancellationToken).ConfigureAwait(false);
                    if (string.Equals(reply.Type, "error", StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(string.IsNullOrWhiteSpace(reply.Message) ? "JDBC 操作失败。" : reply.Message);
                    }

                    if (!string.Equals(reply.Type, "columns", StringComparison.Ordinal) &&
                        !string.Equals(reply.Type, "rows", StringComparison.Ordinal))
                    {
                        return reply;
                    }
                }
            }
            finally
            {
                Forget(request.Id);
            }
        }

        /// <summary>启动查询并等到列信息返回。</summary>
        /// <param name="launch">Java 启动参数。</param>
        /// <param name="target">已解析的 JDBC 连接目标。</param>
        /// <param name="sql">只读 SQL。</param>
        /// <param name="timeoutSeconds">命令超时秒数。</param>
        /// <param name="rowLimit">结果流使用的行数上限。</param>
        /// <param name="connectionName">进度中显示的连接名称。</param>
        /// <param name="progress">可选进度接收器。</param>
        /// <param name="cancellationToken">取消令牌。取消时会通知 Java 中止语句。</param>
        /// <returns>由调用方释放的结果流。</returns>
        public async Task<JdbcResultStream> QueryAsync(
            JdbcBridgeLaunch launch,
            JdbcConnectionTarget target,
            string sql,
            int timeoutSeconds,
            int rowLimit,
            string connectionName,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            if (target == null) throw new ArgumentNullException(nameof(target));
            await EnsureReadyAsync(launch, cancellationToken).ConfigureAwait(false);
            string id = Guid.NewGuid().ToString("N");
            JdbcReplyQueue queue = Register(id);
            JdbcResultStream stream = new JdbcResultStream(this, id, queue, rowLimit, connectionName, progress, cancellationToken);
            try
            {
                Interlocked.Increment(ref _activeQueries);
                stream.Attach();
                progress?.Report(new OperationProgress
                {
                    Stage = OperationStage.Executing,
                    ConnectionName = connectionName,
                    Message = "正在执行只读查询…",
                    IsIndeterminate = true
                });
                Write(new JdbcBridgeRequest
                {
                    Id = id,
                    Op = "query",
                    Url = target.Url,
                    User = target.UserName,
                    Password = target.Password,
                    Driver = target.DriverClass,
                    Jars = target.DriverJars.ToList(),
                    Sql = sql,
                    Catalog = target.Catalog,
                    TimeoutSeconds = timeoutSeconds
                });
                await stream.InitializeAsync(cancellationToken).ConfigureAwait(false);
                return stream;
            }
            catch
            {
                stream.Dispose();
                throw;
            }
        }

        /// <summary>请求 Java 取消指定查询。失败时吞掉异常，避免取消路径再次失败。</summary>
        /// <param name="queryId">查询请求标识。</param>
        public void CancelQuery(string queryId)
        {
            try
            {
                Write(new JdbcBridgeRequest
                {
                    Id = Guid.NewGuid().ToString("N"),
                    Op = "cancel",
                    QueryId = queryId ?? string.Empty
                });
            }
            catch (Exception)
            {
            }
        }

        /// <summary>查询流结束后减少活动计数，以便空闲时可以按新的 Java 路径重启进程。</summary>
        /// <param name="queryId">查询请求标识。</param>
        public void FinishQuery(string queryId)
        {
            Forget(queryId);
            int remaining = Interlocked.Decrement(ref _activeQueries);
            if (remaining < 0)
            {
                Interlocked.Exchange(ref _activeQueries, 0);
            }
        }

        /// <summary>结束转接进程并释放等待中的请求。</summary>
        public void Dispose()
        {
            lock (_gate)
            {
                if (_disposed) return;
                _disposed = true;
            }

            StopProcessCore();
            _lifecycle.Dispose();
        }

        private async Task EnsureReadyAsync(JdbcBridgeLaunch launch, CancellationToken cancellationToken)
        {
            await _lifecycle.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                ThrowIfDisposed();
                string signature = BuildSignature(launch ?? new JdbcBridgeLaunch(string.Empty));
                bool alive;
                bool same;
                int active;
                lock (_gate)
                {
                    alive = _process != null && !_process.HasExited;
                    same = string.Equals(_signature, signature, StringComparison.Ordinal);
                    active = _activeQueries;
                }

                if (alive && (same || active > 0))
                {
                    return;
                }

                StopProcessCore();
                await StartProcessAsync(launch, signature, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _lifecycle.Release();
            }
        }

        private string BuildSignature(JdbcBridgeLaunch launch)
        {
            JavaResolution java = ResolveJava(launch);
            if (!java.Found)
            {
                throw new InvalidOperationException(java.Detail);
            }

            string jar = EnsureBridgeJar();
            return java.Executable + "|" + File.GetLastWriteTimeUtc(jar).Ticks.ToString();
        }

        private string EnsureBridgeJar()
        {
            return JdbcBridgeFiles.EnsureExtracted(_bridgeDirectory);
        }

        private async Task StartProcessAsync(JdbcBridgeLaunch launch, string signature, CancellationToken cancellationToken)
        {
            JavaResolution java = ResolveJava(launch);
            if (!java.Found)
            {
                throw new InvalidOperationException(java.Detail);
            }

            string jar = EnsureBridgeJar();
            Process process = new Process();
            process.StartInfo.FileName = java.Executable;
            process.StartInfo.Arguments = "-Dfile.encoding=UTF-8 -jar " + Quote(jar);
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            if (!process.Start())
            {
                throw new InvalidOperationException("无法启动 Java。");
            }

            UTF8Encoding utf8 = new UTF8Encoding(false);
            StreamWriter input = new StreamWriter(process.StandardInput.BaseStream, utf8) { AutoFlush = true };
            StreamReader output = new StreamReader(process.StandardOutput.BaseStream, utf8);
            StreamReader error = new StreamReader(process.StandardError.BaseStream, utf8);
            int generation;
            lock (_gate)
            {
                _generation++;
                generation = _generation;
                _process = process;
                _input = input;
                _signature = signature;
            }

            Thread reader = new Thread(() => ReadOutput(output, generation))
            {
                IsBackground = true,
                Name = "jdbc-bridge-out"
            };
            reader.Start();
            Thread errors = new Thread(() => ReadErrors(error, generation))
            {
                IsBackground = true,
                Name = "jdbc-bridge-err"
            };
            errors.Start();

            string pingId = Guid.NewGuid().ToString("N");
            JdbcReplyQueue queue = Register(pingId);
            try
            {
                Write(new JdbcBridgeRequest { Id = pingId, Op = "ping" });
                Task<JdbcBridgeReply> pong = queue.DequeueAsync(cancellationToken);
                Task finished = await Task.WhenAny(pong, Task.Delay(30000, cancellationToken)).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                if (finished != pong)
                {
                    throw new InvalidOperationException("JDBC 转接进程没有响应。请确认当前 Java 可以正常启动。");
                }

                JdbcBridgeReply reply = await pong.ConfigureAwait(false);
                if (string.Equals(reply.Type, "error", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(reply.Message) ? "JDBC 转接进程启动失败。" : reply.Message);
                }
            }
            catch
            {
                StopProcessCore();
                throw;
            }
            finally
            {
                Forget(pingId);
            }

            _logger?.Write(LogSeverity.Information, "JDBC 转接进程已启动。");
        }

        private void StopProcessCore()
        {
            Process process;
            StreamWriter input;
            lock (_gate)
            {
                _generation++;
                process = _process;
                input = _input;
                _process = null;
                _input = null;
                _signature = string.Empty;
            }

            FaultAll(new InvalidOperationException("JDBC 转接进程已关闭。"));
            if (process == null) return;
            try
            {
                if (input != null && !process.HasExited)
                {
                    lock (_writeGate)
                    {
                        input.WriteLine(JdbcProtocol.WriteRequest(new JdbcBridgeRequest
                        {
                            Id = Guid.NewGuid().ToString("N"),
                            Op = "shutdown"
                        }));
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (!process.WaitForExit(2000))
                {
                    process.Kill();
                    process.WaitForExit(2000);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                process.Dispose();
            }
            catch (Exception)
            {
            }
        }

        private void ReadOutput(StreamReader reader, int generation)
        {
            try
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    JdbcBridgeReply reply;
                    try
                    {
                        reply = JdbcProtocol.ReadReply(line);
                    }
                    catch (FormatException)
                    {
                        _logger?.Write(LogSeverity.Warning, "忽略了一条无法解析的 JDBC 响应。");
                        continue;
                    }

                    JdbcReplyQueue queue = Find(reply.Id);
                    queue?.Enqueue(reply);
                }
            }
            catch (Exception exception)
            {
                _logger?.Write(LogSeverity.Warning, "读取 JDBC 转接进程输出失败。", exception: exception);
            }
            finally
            {
                bool current;
                lock (_gate)
                {
                    current = generation == _generation;
                    if (current)
                    {
                        _process = null;
                        _input = null;
                    }
                }

                if (current)
                {
                    FaultAll(new InvalidOperationException("JDBC 转接进程已退出。"));
                }
            }
        }

        private void ReadErrors(StreamReader reader, int generation)
        {
            try
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    if (generation != Volatile.Read(ref _generation)) return;
                    if (line.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        line.IndexOf("access_key", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _logger?.Write(LogSeverity.Debug, "JDBC 转接进程输出了已省略的诊断信息。");
                        continue;
                    }

                    if (line.Length > 300) line = line.Substring(0, 300);
                    _logger?.Write(LogSeverity.Debug, "JDBC 转接进程：" + line);
                }
            }
            catch (Exception)
            {
            }
        }

        private void Write(JdbcBridgeRequest request)
        {
            StreamWriter input;
            lock (_gate)
            {
                ThrowIfDisposed();
                input = _input;
            }

            if (input == null)
            {
                throw new InvalidOperationException("JDBC 转接进程尚未启动。");
            }

            string line = JdbcProtocol.WriteRequest(request);
            lock (_writeGate)
            {
                input.WriteLine(line);
            }
        }

        private JdbcReplyQueue Register(string id)
        {
            JdbcReplyQueue queue = new JdbcReplyQueue();
            lock (_gate)
            {
                _queues[id] = queue;
            }

            return queue;
        }

        private JdbcReplyQueue Find(string id)
        {
            lock (_gate)
            {
                _queues.TryGetValue(id ?? string.Empty, out JdbcReplyQueue queue);
                return queue;
            }
        }

        private void Forget(string id)
        {
            lock (_gate)
            {
                _queues.Remove(id ?? string.Empty);
            }
        }

        private void FaultAll(Exception exception)
        {
            List<JdbcReplyQueue> queues;
            lock (_gate)
            {
                queues = new List<JdbcReplyQueue>(_queues.Values);
                _queues.Clear();
            }

            foreach (JdbcReplyQueue queue in queues)
            {
                queue.Fault(exception);
            }
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(JdbcBridgeHost));
        }

        private static string Quote(string value)
        {
            return "\"" + (value ?? string.Empty).Replace("\"", string.Empty) + "\"";
        }

        private static JavaResolution ResolveJava(JdbcBridgeLaunch launch)
        {
            return JavaRuntimeProbe.Resolve(launch == null ? string.Empty : launch.JavaExecutable);
        }

        internal sealed class JdbcReplyQueue
        {
            private readonly ConcurrentQueue<JdbcBridgeReply> _items = new ConcurrentQueue<JdbcBridgeReply>();
            private readonly SemaphoreSlim _signal = new SemaphoreSlim(0);
            private Exception _fault;

            public void Enqueue(JdbcBridgeReply reply)
            {
                _items.Enqueue(reply);
                _signal.Release();
            }

            public void Fault(Exception exception)
            {
                _fault = exception ?? new InvalidOperationException("JDBC 转接进程已中断。");
                _signal.Release();
            }

            public async Task<JdbcBridgeReply> DequeueAsync(CancellationToken cancellationToken)
            {
                while (true)
                {
                    await _signal.WaitAsync(cancellationToken).ConfigureAwait(false);
                    if (_items.TryDequeue(out JdbcBridgeReply reply))
                    {
                        return reply;
                    }

                    if (_fault != null)
                    {
                        throw _fault;
                    }
                }
            }
        }
    }
}
