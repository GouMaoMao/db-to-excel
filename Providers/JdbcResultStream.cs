using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Providers
{
    /// <summary>从 JDBC 转接进程按块读取查询结果，并在达到行数上限后取消语句。</summary>
    /// <remarks>
    /// 本对象拥有对应的查询请求。调用 <see cref="Dispose"/> 会通知 Java 关闭该查询。
    /// 同一实例不能并发读取。行数上限在读取端执行，不改写 SQL。
    /// </remarks>
    internal sealed class JdbcResultStream : IDataResultStream
    {
        private readonly JdbcBridgeHost _host;
        private readonly string _queryId;
        private readonly int _rowLimit;
        private readonly string _connectionName;
        private readonly IProgress<OperationProgress> _progress;
        private readonly CancellationTokenRegistration _cancellationRegistration;
        private readonly Queue<object[]> _pending = new Queue<object[]>();
        private readonly List<string> _columnTypes = new List<string>();
        private readonly JdbcBridgeHost.JdbcReplyQueue _queue;
        private bool _remoteEnded;
        private bool _attached;
        private bool _disposed;
        private bool _limitChecked;
        private bool _cancelRequested;

        /// <summary>创建结果流。列信息要等 <see cref="InitializeAsync"/> 完成后才能读取。</summary>
        /// <param name="host">负责取消和结束计数的进程宿主。</param>
        /// <param name="queryId">查询请求标识。</param>
        /// <param name="queue">该查询的响应队列。</param>
        /// <param name="rowLimit">最大读取行数；小于等于零表示不限制。</param>
        /// <param name="connectionName">进度中显示的连接名称。</param>
        /// <param name="progress">可选进度接收器。</param>
        /// <param name="cancellationToken">取消时通知 Java 中止语句。</param>
        public JdbcResultStream(
            JdbcBridgeHost host,
            string queryId,
            JdbcBridgeHost.JdbcReplyQueue queue,
            int rowLimit,
            string connectionName,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _queryId = queryId ?? string.Empty;
            _queue = queue ?? throw new ArgumentNullException(nameof(queue));
            _rowLimit = rowLimit;
            _connectionName = connectionName ?? string.Empty;
            _progress = progress;
            _cancellationRegistration = cancellationToken.Register(Cancel);
        }

        /// <inheritdoc/>
        public IReadOnlyList<ResultColumn> Columns { get; private set; } = new List<ResultColumn>().AsReadOnly();

        /// <inheritdoc/>
        public long RowsRead { get; private set; }

        /// <inheritdoc/>
        public bool IsCompleted { get; private set; }

        /// <inheritdoc/>
        public bool IsTruncated { get; private set; }

        /// <summary>标记查询已经计入活动数，释放时宿主才会减少计数。</summary>
        public void Attach()
        {
            _attached = true;
        }

        /// <summary>读取列信息。必须在把结果流交给调用方之前完成。</summary>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <returns>表示列信息已经就绪的任务。</returns>
        public async Task InitializeAsync(CancellationToken cancellationToken)
        {
            while (true)
            {
                JdbcBridgeReply reply = await DequeueAsync(cancellationToken).ConfigureAwait(false);
                if (string.Equals(reply.Type, "columns", StringComparison.Ordinal))
                {
                    List<ResultColumn> columns = new List<ResultColumn>();
                    foreach (JdbcColumnInfo column in reply.Columns ?? new List<JdbcColumnInfo>())
                    {
                        string name = string.IsNullOrWhiteSpace(column.Name) ? "column" + (columns.Count + 1) : column.Name;
                        string type = string.IsNullOrWhiteSpace(column.Type) ? "string" : column.Type;
                        _columnTypes.Add(type);
                        columns.Add(new ResultColumn(name, JdbcProtocol.ToClrType(type)));
                    }

                    Columns = columns.AsReadOnly();
                    return;
                }

                if (string.Equals(reply.Type, "end", StringComparison.Ordinal))
                {
                    Columns = new List<ResultColumn>().AsReadOnly();
                    _remoteEnded = true;
                    IsCompleted = true;
                    return;
                }

                if (string.Equals(reply.Type, "error", StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(reply.Message) ? "JDBC 查询失败。" : reply.Message);
                }
            }
        }

        /// <inheritdoc/>
        public async Task<ResultBlock> ReadBlockAsync(int maximumRows, CancellationToken cancellationToken)
        {
            ThrowIfDisposed();
            if (maximumRows <= 0) throw new ArgumentOutOfRangeException(nameof(maximumRows));
            if (IsCompleted) return new ResultBlock(new object[0, Columns.Count], 0, true);

            int remaining = _rowLimit > 0
                ? (int)Math.Min(maximumRows, Math.Max(0L, _rowLimit - RowsRead))
                : maximumRows;
            if (remaining == 0)
            {
                await CheckLimitAsync(cancellationToken).ConfigureAwait(false);
                return new ResultBlock(new object[0, Columns.Count], 0, true);
            }

            object[,] values = new object[remaining, Columns.Count];
            int rowCount = 0;
            while (rowCount < remaining && await TryTakeRowAsync(cancellationToken).ConfigureAwait(false))
            {
                object[] row = _pending.Dequeue();
                for (int column = 0; column < Columns.Count; column++)
                {
                    values[rowCount, column] = column < row.Length ? row[column] : null;
                }

                rowCount++;
                RowsRead++;
            }

            if (rowCount < remaining || _remoteEnded && _pending.Count == 0)
            {
                IsCompleted = true;
            }
            else if (_rowLimit > 0 && RowsRead >= _rowLimit)
            {
                await CheckLimitAsync(cancellationToken).ConfigureAwait(false);
            }

            object[,] result = rowCount == remaining ? values : CopyRows(values, rowCount, Columns.Count);
            _progress?.Report(new OperationProgress
            {
                Stage = OperationStage.Reading,
                ConnectionName = _connectionName,
                Message = IsTruncated
                    ? string.Format(
                        "已读取 {0:N0} 行，达到上限 {1:N0} 行，结果已截断。可在设置中修改上限行数。",
                        RowsRead,
                        _rowLimit)
                    : "正在读取结果…",
                RowsRead = RowsRead,
                IsIndeterminate = true,
                IsTruncated = IsTruncated
            });
            return new ResultBlock(result, rowCount, IsCompleted);
        }

        /// <inheritdoc/>
        public void Cancel()
        {
            _cancelRequested = true;
            _host.CancelQuery(_queryId);
        }

        /// <summary>取消尚未结束的查询，并通知宿主减少活动计数。</summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                _cancellationRegistration.Dispose();
            }
            catch (Exception)
            {
            }

            if (!_remoteEnded)
            {
                Cancel();
            }

            if (_attached)
            {
                _attached = false;
                _host.FinishQuery(_queryId);
            }
        }

        private async Task<bool> TryTakeRowAsync(CancellationToken cancellationToken)
        {
            if (_pending.Count > 0) return true;
            if (_remoteEnded) return false;
            while (_pending.Count == 0 && !_remoteEnded)
            {
                JdbcBridgeReply reply = await DequeueAsync(cancellationToken).ConfigureAwait(false);
                if (string.Equals(reply.Type, "rows", StringComparison.Ordinal))
                {
                    foreach (object[] row in reply.Rows ?? new List<object[]>())
                    {
                        object[] converted = new object[Columns.Count];
                        for (int index = 0; index < Columns.Count; index++)
                        {
                            object raw = row != null && index < row.Length ? row[index] : null;
                            converted[index] = JdbcProtocol.ConvertCell(raw, index < _columnTypes.Count ? _columnTypes[index] : "string");
                        }

                        _pending.Enqueue(converted);
                    }
                }
                else if (string.Equals(reply.Type, "end", StringComparison.Ordinal))
                {
                    _remoteEnded = true;
                }
                else if (string.Equals(reply.Type, "error", StringComparison.Ordinal))
                {
                    if (_cancelRequested)
                    {
                        _remoteEnded = true;
                    }
                    else
                    {
                        throw new InvalidOperationException(string.IsNullOrWhiteSpace(reply.Message) ? "JDBC 查询失败。" : reply.Message);
                    }
                }
            }

            return _pending.Count > 0;
        }

        private async Task CheckLimitAsync(CancellationToken cancellationToken)
        {
            if (_limitChecked) return;
            _limitChecked = true;
            if (await TryTakeRowAsync(cancellationToken).ConfigureAwait(false))
            {
                IsTruncated = true;
                Cancel();
                _pending.Clear();
            }

            IsCompleted = true;
            _remoteEnded = true;
        }

        private Task<JdbcBridgeReply> DequeueAsync(CancellationToken cancellationToken)
        {
            return _queue.DequeueAsync(cancellationToken);
        }

        private static object[,] CopyRows(object[,] source, int rows, int columns)
        {
            object[,] target = new object[rows, columns];
            for (int row = 0; row < rows; row++)
            {
                for (int column = 0; column < columns; column++)
                {
                    target[row, column] = source[row, column];
                }
            }

            return target;
        }

        private void ThrowIfDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(JdbcResultStream));
        }
    }
}
