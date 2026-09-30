using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Providers
{
    /// <summary>
    /// 包装 ADO.NET 连接、命令和读取器，提供受行数限制的分块异步读取。
    /// </summary>
    /// <remarks>
    /// 此类型拥有传入的连接、命令和读取器。调用 <see cref="Dispose"/> 会按读取器、命令、连接的顺序释放资源。
    /// 每读完一块报告已读行数。总行数在读完前未知，因此不报告完成百分比。
    /// 实例不是线程安全的，同一时间只能进行一次读取。
    /// </remarks>
    internal sealed class DatabaseResultStream : IDataResultStream
    {
        private readonly DbConnection _connection;
        private readonly DbCommand _command;
        private readonly DbDataReader _reader;
        private readonly int _rowLimit;
        private readonly IProgress<OperationProgress> _progress;
        private readonly CancellationTokenRegistration _cancellationRegistration;
        private bool _disposed;
        private bool _limitChecked;

        /// <summary>创建并接管数据库查询资源。</summary>
        /// <param name="connection">已打开的连接。</param>
        /// <param name="command">正在执行的命令。</param>
        /// <param name="reader">命令返回的数据读取器。</param>
        /// <param name="rowLimit">最大读取行数；小于等于零表示不限制。</param>
        /// <param name="progress">可选进度接收器。</param>
        /// <param name="cancellationToken">取消时会调用数据库命令的取消方法。</param>
        public DatabaseResultStream(
            DbConnection connection,
            DbCommand command,
            DbDataReader reader,
            int rowLimit,
            IProgress<OperationProgress> progress,
            CancellationToken cancellationToken)
        {
            _connection = connection ?? throw new ArgumentNullException(nameof(connection));
            _command = command ?? throw new ArgumentNullException(nameof(command));
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _rowLimit = rowLimit;
            _progress = progress;
            Columns = Enumerable.Range(0, reader.FieldCount)
                .Select(index => new ResultColumn(reader.GetName(index), reader.GetFieldType(index)))
                .ToList()
                .AsReadOnly();
            _cancellationRegistration = cancellationToken.Register(Cancel);
        }

        /// <inheritdoc/>
        public IReadOnlyList<ResultColumn> Columns { get; }
        /// <inheritdoc/>
        public long RowsRead { get; private set; }
        /// <inheritdoc/>
        public bool IsCompleted { get; private set; }
        /// <inheritdoc/>
        public bool IsTruncated { get; private set; }

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
            while (rowCount < remaining && await _reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                for (int column = 0; column < Columns.Count; column++)
                {
                    object value = _reader.GetValue(column);
                    values[rowCount, column] = value == DBNull.Value ? null : value;
                }

                rowCount++;
                RowsRead++;
            }

            if (rowCount < remaining)
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
                Message = IsTruncated
                    ? string.Format(
                        "已读取 {0:N0} 行，达到上限 {1:N0} 行，结果已截断。可在设置中修改上限行数。",
                        RowsRead,
                        _rowLimit)
                    : string.Format("正在读取 {0:N0} 行", RowsRead),
                RowsRead = RowsRead,
                IsIndeterminate = true,
                IsTruncated = IsTruncated,
                Severity = IsTruncated ? LogSeverity.Warning : LogSeverity.Information
            });
            return new ResultBlock(result, rowCount, IsCompleted);
        }

        /// <inheritdoc/>
        public void Cancel()
        {
            try
            {
                _command.Cancel();
            }
            catch (Exception)
            {
            }
        }

        /// <summary>释放取消注册、读取器、命令和连接。</summary>
        /// <remarks>主动截断后驱动可能把取消当成语句失败；清理阶段的异常不再冒泡，避免成功读取被记成用户取消。</remarks>
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
            DisposeQuietly(_reader);
            DisposeQuietly(_command);
            DisposeQuietly(_connection);
        }

        private static void DisposeQuietly(IDisposable disposable)
        {
            if (disposable == null) return;
            try
            {
                disposable.Dispose();
            }
            catch (Exception)
            {
            }
        }

        private async Task CheckLimitAsync(CancellationToken cancellationToken)
        {
            if (_limitChecked) return;
            _limitChecked = true;
            if (await _reader.ReadAsync(cancellationToken).ConfigureAwait(false))
            {
                IsTruncated = true;
                Cancel();
            }
            IsCompleted = true;
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
            if (_disposed) throw new ObjectDisposedException(nameof(DatabaseResultStream));
        }
    }
}
