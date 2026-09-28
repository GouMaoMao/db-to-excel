using System;
using DB2Sheet.Excel;
using DB2Sheet.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Excel
{
    [TestClass]
    public sealed class SqlSheetTaskParserTests
    {
        private readonly SqlSheetTaskParser _parser = new SqlSheetTaskParser();

        [TestMethod]
        public void Parse_CreatesOneTaskPerPopulatedColumn()
        {
            object[,] cells =
            {
                { " Customers ", "Orders" },
                { "SELECT id", "SELECT number" },
                { "FROM customers", "FROM orders" }
            };

            var tasks = _parser.Parse(cells, 3);

            Assert.AreEqual(2, tasks.Count);
            AssertTask(tasks[0], "column-3", "Customers", "SELECT id" + Environment.NewLine + "FROM customers", 3);
            AssertTask(tasks[1], "column-4", "Orders", "SELECT number" + Environment.NewLine + "FROM orders", 4);
        }

        [TestMethod]
        public void Parse_SkipsEmptyTargetsEmptyQueriesAndBlankFragments()
        {
            object[,] cells =
            {
                { null, "EmptyQuery", "Reports" },
                { "SELECT 1", "  ", "SELECT id" },
                { null, null, "" },
                { null, null, "FROM reports" }
            };

            var tasks = _parser.Parse(cells, 1);

            Assert.AreEqual(1, tasks.Count);
            AssertTask(tasks[0], "column-3", "Reports", "SELECT id" + Environment.NewLine + "FROM reports", 3);
        }

        [TestMethod]
        public void Parse_SupportsNonZeroArrayBoundsAndPreservesSourceColumn()
        {
            object[,] cells = (object[,])Array.CreateInstance(typeof(object), new[] { 3, 1 }, new[] { 1, 5 });
            cells[1, 5] = "Summary";
            cells[2, 5] = "SELECT";
            cells[3, 5] = " 1";

            var tasks = _parser.Parse(cells, 7);

            Assert.AreEqual(1, tasks.Count);
            AssertTask(tasks[0], "column-7", "Summary", "SELECT" + Environment.NewLine + " 1", 7);
        }

        [TestMethod]
        public void Parse_RejectsInvalidArguments()
        {
            Assert.ThrowsException<ArgumentNullException>(() => _parser.Parse(null, 1));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => _parser.Parse(new object[2, 1], 0));
        }

        private static void AssertTask(
            RefreshTaskDefinition task,
            string expectedId,
            string expectedTarget,
            string expectedQuery,
            int expectedSourceColumn)
        {
            Assert.AreEqual(expectedId, task.Id);
            Assert.AreEqual(expectedTarget, task.TargetSheetName);
            Assert.AreEqual(expectedQuery, task.QueryText);
            Assert.AreEqual(expectedSourceColumn, task.SourceColumn);
        }
    }
}
