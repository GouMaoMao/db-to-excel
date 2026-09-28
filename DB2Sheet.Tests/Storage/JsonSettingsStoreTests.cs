using System;
using System.Collections.Generic;
using System.IO;
using DB2Sheet.Infrastructure;
using DB2Sheet.Services;
using DB2Sheet.Storage;
using DB2Sheet.Tests.TestInfrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Storage
{
    [TestClass]
    public sealed class JsonSettingsStoreTests
    {
        [TestMethod]
        public void Constructor_UsesDefaultsWhenFileDoesNotExist()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                JsonSettingsStore store = CreateStore(paths);

                Assert.AreEqual(100, store.Get(CoreSettings.MaxPreviewRows));
                Assert.AreEqual(0, store.GetRawValues().Count);
                Assert.IsFalse(File.Exists(paths.SettingsFile));
            }
        }

        [TestMethod]
        public void Set_PersistsValueAndRaisesChangedOnce()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                JsonSettingsStore store = CreateStore(paths);
                int changedCount = 0;
                store.Changed += (sender, args) => changedCount++;

                store.Set(CoreSettings.MaxPreviewRows, 250);

                Assert.AreEqual(1, changedCount);
                Assert.AreEqual(250, store.Get(CoreSettings.MaxPreviewRows));
                Assert.AreEqual(250, CreateStore(paths).Get(CoreSettings.MaxPreviewRows));
                Assert.IsTrue(File.Exists(paths.SettingsFile));
            }
        }

        [TestMethod]
        public void Set_RejectsInvalidValueWithoutChangingStore()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                JsonSettingsStore store = CreateStore(paths);
                int changedCount = 0;
                store.Changed += (sender, args) => changedCount++;

                Assert.ThrowsException<ArgumentOutOfRangeException>(() => store.Set(CoreSettings.MaxPreviewRows, 0));

                Assert.AreEqual(0, changedCount);
                Assert.AreEqual(100, store.Get(CoreSettings.MaxPreviewRows));
                Assert.IsFalse(File.Exists(paths.SettingsFile));
            }
        }

        [TestMethod]
        public void SetMany_PersistsAllValuesAndRaisesSingleEvent()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                JsonSettingsStore store = CreateStore(paths);
                int changedCount = 0;
                store.Changed += (sender, args) => changedCount++;

                store.SetMany(new Dictionary<string, object>
                {
                    [CoreSettings.MaxPreviewRows.Key] = 500,
                    [CoreSettings.QueryTimeoutSeconds.Key] = 60
                });

                JsonSettingsStore reloaded = CreateStore(paths);
                Assert.AreEqual(1, changedCount);
                Assert.AreEqual(500, reloaded.Get(CoreSettings.MaxPreviewRows));
                Assert.AreEqual(60, reloaded.Get(CoreSettings.QueryTimeoutSeconds));
            }
        }

        [TestMethod]
        public void SetMany_ValidatesEntireBatchBeforeSaving()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                JsonSettingsStore store = CreateStore(paths);

                Assert.ThrowsException<ArgumentException>(() => store.SetMany(new Dictionary<string, object>
                {
                    [CoreSettings.MaxPreviewRows.Key] = 500,
                    ["unknown.setting"] = "value"
                }));

                Assert.AreEqual(100, store.Get(CoreSettings.MaxPreviewRows));
                Assert.IsFalse(File.Exists(paths.SettingsFile));
            }
        }

        [TestMethod]
        public void Save_PreservesUnknownKeysAndRawValuesAreCopied()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                File.WriteAllText(paths.SettingsFile,
                    "{\"Version\":1,\"Values\":{\"future.option\":\"kept\",\"query.maxPreviewRows\":\"150\"}}");
                JsonSettingsStore store = CreateStore(paths);
                Dictionary<string, string> raw = (Dictionary<string, string>)store.GetRawValues();
                raw["future.option"] = "changed-outside";

                store.Set(CoreSettings.QueryTimeoutSeconds, 45);

                JsonSettingsStore reloaded = CreateStore(paths);
                Assert.AreEqual("kept", reloaded.GetRawValues()["future.option"]);
                Assert.AreEqual(150, reloaded.Get(CoreSettings.MaxPreviewRows));
                Assert.AreEqual(45, reloaded.Get(CoreSettings.QueryTimeoutSeconds));
            }
        }

        [TestMethod]
        public void Get_FallsBackToDefaultForInvalidPersistedValue()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                File.WriteAllText(paths.SettingsFile,
                    "{\"Version\":1,\"Values\":{\"query.maxPreviewRows\":\"0\"}}");

                JsonSettingsStore store = CreateStore(paths);

                Assert.AreEqual(100, store.Get(CoreSettings.MaxPreviewRows));
                Assert.AreEqual("0", store.GetRawValues()[CoreSettings.MaxPreviewRows.Key]);
            }
        }

        [TestMethod]
        public void Constructor_BacksUpCorruptFileAndUsesDefaults()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                File.WriteAllText(paths.SettingsFile, "not-json");

                JsonSettingsStore store = CreateStore(paths);

                Assert.AreEqual(100, store.Get(CoreSettings.MaxPreviewRows));
                Assert.IsFalse(File.Exists(paths.SettingsFile));
                Assert.AreEqual(1, Directory.GetFiles(temporary.Path, "settings.json.corrupt-*").Length);
            }
        }

        private static JsonSettingsStore CreateStore(ApplicationPaths paths)
        {
            return new JsonSettingsStore(paths, SettingsRegistry.CreateDefault());
        }
    }
}
