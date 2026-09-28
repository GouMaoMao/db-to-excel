using System.Linq;
using DB2Sheet.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Services
{
    [TestClass]
    public sealed class SqlReadOnlyValidatorTests
    {
        private readonly SqlReadOnlyValidator _validator = new SqlReadOnlyValidator();

        [TestMethod]
        public void Validate_AcceptsSelectAndReadOnlyCte()
        {
            Assert.AreEqual(0, _validator.Validate("SELECT id, name FROM users", "sqlserver").Count);
            Assert.AreEqual(0, _validator.Validate("WITH active AS (SELECT id FROM users) SELECT id FROM active;", "postgresql").Count);
        }

        [TestMethod]
        public void Validate_RejectsEmptyAndNonQueryStatements()
        {
            Assert.IsTrue(_validator.Validate("  ", "sqlserver").Any(error => error.Contains("不能为空")));
            Assert.IsTrue(_validator.Validate("UPDATE users SET active = 1", "sqlserver").Any(error => error.Contains("UPDATE")));
            Assert.IsTrue(_validator.Validate("EXEC report", "sqlserver").Any(error => error.Contains("EXEC")));
        }

        [TestMethod]
        public void Validate_RejectsMultipleStatementsButAllowsTrailingSemicolon()
        {
            Assert.AreEqual(0, _validator.Validate("SELECT 1;", "sqlserver").Count);
            Assert.IsTrue(_validator.Validate("SELECT 1; SELECT 2", "sqlserver").Any(error => error.Contains("单条")));
        }

        [TestMethod]
        public void Validate_IgnoresForbiddenWordsInCommentsAndLiterals()
        {
            string sql = "SELECT 'DELETE FROM users' AS text_value /* UPDATE users */ -- DROP table\r\nFROM reports";

            Assert.AreEqual(0, _validator.Validate(sql, "sqlserver").Count);
        }

        [TestMethod]
        public void Validate_RejectsSelectIntoAndLockingClauses()
        {
            Assert.IsTrue(_validator.Validate("SELECT id INTO archived FROM users", "sqlserver").Any(error => error.Contains("INTO")));
            Assert.IsTrue(_validator.Validate("SELECT id FROM users FOR UPDATE", "postgresql").Any(error => error.Contains("锁定")));
            Assert.IsTrue(_validator.Validate("SELECT id FROM users FOR SHARE", "postgresql").Any(error => error.Contains("锁定")));
        }

        [TestMethod]
        public void Validate_AppliesPostgreSqlReturningRuleOnlyToPostgreSql()
        {
            Assert.IsTrue(_validator.Validate("SELECT returning FROM audit_log", "postgresql").Any(error => error.Contains("RETURNING")));
            Assert.AreEqual(0, _validator.Validate("SELECT returning FROM audit_log", "mysql").Count);
        }

        [TestMethod]
        public void Validate_ReportsUnclosedCommentOrLiteral()
        {
            Assert.IsTrue(_validator.Validate("SELECT 1 /* open", "sqlserver").Any(error => error.Contains("注释未闭合")));
            Assert.IsTrue(_validator.Validate("SELECT 'open", "sqlserver").Any(error => error.Contains("未闭合")));
        }

        [TestMethod]
        public void Validate_AllowsQuotedIdentifiersContainingForbiddenWords()
        {
            Assert.AreEqual(0, _validator.Validate("SELECT [delete], `update`, \"drop\" FROM audit_log", "sqlite").Count);
        }
    }
}
