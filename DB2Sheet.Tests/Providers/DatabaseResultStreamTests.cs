using System;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Threading;
using System.Threading.Tasks;
using DB2Sheet.Models;
using DB2Sheet.Providers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Providers
{
    [TestClass]
    public sealed class DatabaseResultStreamTests
    {
        [TestMethod]
        public async Task ReadBlockAsync_ReachingRowLimitMarksTruncatedAndCancelsCommand()
        {
            DataTable table = new DataTable();
            table.Columns.Add("id", typeof(int));
            table.Rows.Add(1);
            table.Rows.Add(2);
            table.Rows.Add(3);

            using (DbDataReader reader = table.CreateDataReader())
            using (FakeDbConnection connection = new FakeDbConnection())
            using (TrackingDbCommand command = new TrackingDbCommand())
            using (DatabaseResultStream stream = new DatabaseResultStream(
                connection,
                command,
                reader,
                2,
                null,
                CancellationToken.None))
            {
                ResultBlock block = await stream.ReadBlockAsync(2, CancellationToken.None);

                Assert.AreEqual(2, block.RowCount);
                Assert.IsTrue(stream.IsTruncated);
                Assert.IsTrue(stream.IsCompleted);
                Assert.IsTrue(command.CancelCalled);
            }
        }

        private sealed class TrackingDbCommand : DbCommand
        {
            private readonly SqlCommand _inner = new SqlCommand();

            public bool CancelCalled { get; private set; }

            public override string CommandText { get; set; }
            public override int CommandTimeout { get; set; }
            public override CommandType CommandType { get; set; }
            public override bool DesignTimeVisible { get; set; }
            public override UpdateRowSource UpdatedRowSource { get; set; }

            protected override DbConnection DbConnection { get; set; }
            protected override DbParameterCollection DbParameterCollection => _inner.Parameters;
            protected override DbTransaction DbTransaction { get; set; }

            public override void Cancel()
            {
                CancelCalled = true;
            }

            public override int ExecuteNonQuery()
            {
                throw new NotSupportedException();
            }

            public override object ExecuteScalar()
            {
                throw new NotSupportedException();
            }

            public override void Prepare()
            {
            }

            protected override DbParameter CreateDbParameter()
            {
                return new SqlParameter();
            }

            protected override DbDataReader ExecuteDbDataReader(CommandBehavior behavior)
            {
                throw new NotSupportedException();
            }
        }

        private sealed class FakeDbConnection : DbConnection
        {
            public override string ConnectionString { get; set; }
            public override string Database => "test";
            public override string DataSource => "test";
            public override string ServerVersion => "1.0";
            public override ConnectionState State => ConnectionState.Open;

            public override void ChangeDatabase(string databaseName)
            {
            }

            public override void Close()
            {
            }

            public override void Open()
            {
            }

            protected override DbTransaction BeginDbTransaction(IsolationLevel isolationLevel)
            {
                throw new NotSupportedException();
            }

            protected override DbCommand CreateDbCommand()
            {
                throw new NotSupportedException();
            }
        }
    }
}
