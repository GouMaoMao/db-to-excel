using System;
using System.Collections.Generic;
using System.IO;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;
using DB2Sheet.Storage;
using DB2Sheet.Tests.TestInfrastructure;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DB2Sheet.Tests.Storage
{
    [TestClass]
    public sealed class QueryProfileRepositoryTests
    {
        [TestMethod]
        public void Save_AssignsIdUpdatesTimestampAndPersistsProfile()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                QueryProfileRepository repository = new QueryProfileRepository(paths);
                QueryProfile profile = CreateProfile("Customer report");
                profile.Id = null;
                profile.UpdatedUtc = DateTime.UtcNow.AddDays(-1);
                DateTime previousUpdated = profile.UpdatedUtc;

                repository.Save(profile);

                Assert.IsFalse(string.IsNullOrWhiteSpace(profile.Id));
                QueryProfile reloaded = new QueryProfileRepository(paths).GetById(profile.Id);
                Assert.AreEqual("Customer report", reloaded.Name);
                Assert.AreEqual("SELECT * FROM customers", reloaded.QueryText);
                Assert.IsTrue(reloaded.UpdatedUtc > previousUpdated);
            }
        }

        [TestMethod]
        public void Save_AndGetReturnCopiesIsolatedFromExternalChanges()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                QueryProfileRepository repository = CreateRepository(temporary);
                QueryProfile profile = CreateProfile("Report");
                repository.Save(profile);

                profile.Name = "Changed outside";
                profile.ProviderOptions["schema"] = "changed";
                QueryProfile first = repository.GetById(profile.Id);
                first.Name = "Changed result";
                first.ProviderOptions["schema"] = "changed-result";
                QueryProfile second = repository.GetById(profile.Id);

                Assert.AreEqual("Report", second.Name);
                Assert.AreEqual("public", second.ProviderOptions["schema"]);
            }
        }

        [TestMethod]
        public void GetAll_SortsByNameAndReturnsIndependentCopies()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                QueryProfileRepository repository = CreateRepository(temporary);
                repository.Save(CreateProfile("Zulu"));
                repository.Save(CreateProfile("Alpha"));

                var profiles = repository.GetAll();
                profiles[0].Name = "Changed";

                Assert.AreEqual("Alpha", repository.GetAll()[0].Name);
                Assert.AreEqual("Zulu", repository.GetAll()[1].Name);
            }
        }

        [TestMethod]
        public void Save_OverwritesExistingProfileWithoutAddingDuplicate()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                QueryProfileRepository repository = new QueryProfileRepository(paths);
                QueryProfile profile = CreateProfile("Original");
                DateTime created = profile.CreatedUtc;
                repository.Save(profile);
                profile.Name = "Updated";
                profile.QueryText = "SELECT 2";

                repository.Save(profile);

                QueryProfileRepository reloaded = new QueryProfileRepository(paths);
                Assert.AreEqual(1, reloaded.GetAll().Count);
                Assert.AreEqual("Updated", reloaded.GetById(profile.Id).Name);
                Assert.AreEqual("SELECT 2", reloaded.GetById(profile.Id).QueryText);
                Assert.IsTrue((created - reloaded.GetById(profile.Id).CreatedUtc).Duration() < TimeSpan.FromSeconds(1));
            }
        }

        [TestMethod]
        public void GetById_IsCaseInsensitiveAndDeletePersistsRemoval()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                QueryProfileRepository repository = new QueryProfileRepository(paths);
                QueryProfile profile = CreateProfile("Report");
                profile.Id = "QueryABC";
                repository.Save(profile);

                Assert.IsNotNull(repository.GetById("queryabc"));
                repository.Delete("QUERYABC");
                repository.Delete("missing");

                Assert.IsNull(repository.GetById(profile.Id));
                Assert.AreEqual(0, new QueryProfileRepository(paths).GetAll().Count);
            }
        }

        [TestMethod]
        public void Save_ValidatesNameAndNormalizesNullProviderOptions()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                QueryProfileRepository repository = CreateRepository(temporary);
                Assert.ThrowsException<ArgumentNullException>(() => repository.Save(null));
                Assert.ThrowsException<ArgumentException>(() => repository.Save(CreateProfile(" ")));
                QueryProfile profile = CreateProfile("Report");
                profile.ProviderOptions = null;

                repository.Save(profile);

                Assert.AreEqual(0, repository.GetById(profile.Id).ProviderOptions.Count);
            }
        }

        [TestMethod]
        public void Constructor_BacksUpCorruptFileAndStartsEmpty()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                File.WriteAllText(paths.QueriesFile, "not-json");

                QueryProfileRepository repository = new QueryProfileRepository(paths);

                Assert.AreEqual(0, repository.GetAll().Count);
                Assert.IsFalse(File.Exists(paths.QueriesFile));
                Assert.AreEqual(1, Directory.GetFiles(temporary.Path, "queries.json.corrupt-*").Length);
            }
        }

        [TestMethod]
        public void Save_PersistsEachProfileToDedicatedJsonFile()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                QueryProfileRepository repository = new QueryProfileRepository(paths);
                repository.Save(CreateProfile("One"));
                repository.Save(CreateProfile("Two"));

                Assert.AreEqual(2, repository.GetAll().Count);
                Assert.AreEqual(2, Directory.GetFiles(paths.QueryProfilesDirectory, "*.json").Length);
            }
        }

        [TestMethod]
        public void Constructor_MigratesLegacyQueriesFileToProfileDirectory()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                string legacyJson =
                    "{\"Version\":1,\"Profiles\":[{\"Id\":\"legacy-1\",\"Name\":\"Legacy\",\"ProviderId\":\"postgresql\",\"ConnectionProfileId\":\"connection-1\",\"QueryText\":\"SELECT 1\",\"TargetSheetName\":\"Sheet1\",\"Description\":\"\",\"ProviderOptions\":{}}]}";
                File.WriteAllText(paths.QueriesFile, legacyJson);

                QueryProfileRepository repository = new QueryProfileRepository(paths);

                Assert.AreEqual(1, repository.GetAll().Count);
                Assert.IsFalse(File.Exists(paths.QueriesFile));
                Assert.AreEqual(1, Directory.GetFiles(paths.QueryProfilesDirectory, "*.json").Length);
                Assert.AreEqual(1, Directory.GetFiles(temporary.Path, "queries.json.migrated-*").Length);
            }
        }

        private static QueryProfileRepository CreateRepository(TemporaryDirectory temporary)
        {
            return new QueryProfileRepository(new ApplicationPaths(temporary.Path));
        }

        private static QueryProfile CreateProfile(string name)
        {
            return new QueryProfile
            {
                Name = name,
                ProviderId = "postgresql",
                ConnectionProfileId = "connection-1",
                QueryText = "SELECT * FROM customers",
                TargetSheetName = "Customers",
                Description = "Test query",
                ProviderOptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["schema"] = "public"
                }
            };
        }
    }
}
