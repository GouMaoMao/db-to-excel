using System;
using System.Linq;
using DB2Sheet.Models;
using DB2Sheet.Services;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Services
{
    [TestClass]
    public sealed class SettingsRegistryTests
    {
        [TestMethod]
        public void CreateDefault_RegistersAllCoreSettings()
        {
            SettingsRegistry registry = SettingsRegistry.CreateDefault();

            Assert.AreEqual(7, registry.GetAll().Count);
            Assert.AreSame(CoreSettings.MaxPreviewRows, registry.GetByKey("QUERY.MAXPREVIEWROWS"));
            Assert.AreSame(CoreSettings.MaxExportRows, registry.GetByKey("excel.maxExportRows"));
            Assert.IsNull(registry.GetByKey("unknown"));
            Assert.IsNull(registry.GetByKey(null));
        }

        [TestMethod]
        public void GetAll_SortsByCategoryThenDisplayName()
        {
            SettingsRegistry registry = SettingsRegistry.CreateDefault();
            string[] actual = registry.GetAll()
                .Select(definition => definition.Category + "/" + definition.DisplayName)
                .ToArray();
            string[] expected = actual.OrderBy(value => value).ToArray();

            CollectionAssert.AreEqual(expected, actual);
        }

        [TestMethod]
        public void Register_RejectsNullAndDuplicateKeysIgnoringCase()
        {
            SettingsRegistry registry = new SettingsRegistry();
            SettingDefinition<string> definition = CreateStringDefinition("feature.name");
            registry.Register(definition);

            Assert.ThrowsException<ArgumentNullException>(() => registry.Register(null));
            Assert.ThrowsException<InvalidOperationException>(() => registry.Register(CreateStringDefinition("FEATURE.NAME")));
        }

        [TestMethod]
        public void IntegerSetting_ParsesInvariantValueAndEnforcesRange()
        {
            Assert.IsTrue(CoreSettings.MaxPreviewRows.TryParse("250", out object parsed));
            Assert.AreEqual(250, parsed);
            Assert.IsFalse(CoreSettings.MaxPreviewRows.TryParse("0", out _));
            Assert.IsFalse(CoreSettings.MaxPreviewRows.TryParse("100001", out _));
            Assert.IsFalse(CoreSettings.MaxPreviewRows.TryParse("invalid", out _));
        }

        [TestMethod]
        public void BatchMode_ParsesCaseInsensitivelyAndRejectsUnknownValue()
        {
            Assert.IsTrue(CoreSettings.BatchMode.TryParse("parallel", out object parsed));
            Assert.AreEqual(BatchExecutionMode.Parallel, parsed);
            Assert.IsFalse(CoreSettings.BatchMode.TryParse("concurrent", out _));
        }

        [TestMethod]
        public void LogLevel_ValidatesAllowedValues()
        {
            Assert.IsTrue(CoreSettings.LogLevel.TryParse("Debug", out object parsed));
            Assert.AreEqual("Debug", parsed);
            Assert.IsFalse(CoreSettings.LogLevel.TryParse("Verbose", out _));
            Assert.IsFalse(CoreSettings.LogLevel.TryParse("", out _));
        }

        [TestMethod]
        public void SettingDefinition_SerializesInvariantValueAndRejectsWrongType()
        {
            Assert.AreEqual("300", CoreSettings.QueryTimeoutSeconds.Serialize(300));
            Assert.ThrowsException<ArgumentException>(() => CoreSettings.QueryTimeoutSeconds.Serialize("300"));
        }

        private static SettingDefinition<string> CreateStringDefinition(string key)
        {
            return new SettingDefinition<string>(
                key,
                "功能名称",
                "扩展",
                "default",
                value => Tuple.Create(true, value));
        }
    }
}
