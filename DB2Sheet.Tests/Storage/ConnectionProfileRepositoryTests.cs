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
    public sealed class ConnectionProfileRepositoryTests
    {
        [TestMethod]
        public void Save_AssignsIdPersistsProfileAndRaisesChanged()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                ConnectionProfileRepository repository = new ConnectionProfileRepository(paths);
                ConnectionProfilesChangedEventArgs changed = null;
                repository.Changed += (sender, args) => changed = args;
                ConnectionProfile profile = CreateProfile("Primary", "sqlserver");
                profile.Id = null;

                repository.Save(profile);

                Assert.IsFalse(string.IsNullOrWhiteSpace(profile.Id));
                Assert.AreEqual(profile.Id, changed.ProfileId);
                ConnectionProfileSnapshot reloaded = new ConnectionProfileRepository(paths).GetById(profile.Id);
                Assert.AreEqual("Primary", reloaded.Name);
                Assert.AreEqual("sqlserver", reloaded.ProviderId);
                Assert.AreEqual("localhost", reloaded.GetValue("server"));
            }
        }

        [TestMethod]
        public void Save_CopiesInputAndGetAllSortsByName()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ConnectionProfileRepository repository = CreateRepository(temporary);
                ConnectionProfile zulu = CreateProfile("Zulu", "mysql");
                ConnectionProfile alpha = CreateProfile("Alpha", "postgresql");
                repository.Save(zulu);
                repository.Save(alpha);

                zulu.Name = "Changed";
                zulu.Parameters["server"] = "changed";
                var profiles = repository.GetAll();

                Assert.AreEqual(2, profiles.Count);
                Assert.AreEqual("Alpha", profiles[0].Name);
                Assert.AreEqual("Zulu", profiles[1].Name);
                Assert.AreEqual("localhost", profiles[1].GetValue("server"));
            }
        }

        [TestMethod]
        public void Save_OverwritesExistingProfileWithoutAddingDuplicate()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                ConnectionProfileRepository repository = new ConnectionProfileRepository(paths);
                ConnectionProfile profile = CreateProfile("Original", "sqlite");
                repository.Save(profile);
                profile.Name = "Updated";
                profile.Parameters["file"] = "database.db";

                repository.Save(profile);

                ConnectionProfileRepository reloaded = new ConnectionProfileRepository(paths);
                Assert.AreEqual(1, reloaded.GetAll().Count);
                Assert.AreEqual("Updated", reloaded.GetById(profile.Id).Name);
                Assert.AreEqual("database.db", reloaded.GetById(profile.Id).GetValue("file"));
            }
        }

        [TestMethod]
        public void GetById_IsCaseInsensitiveAndReturnsNullForMissingId()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ConnectionProfileRepository repository = CreateRepository(temporary);
                ConnectionProfile profile = CreateProfile("Primary", "sqlserver");
                profile.Id = "ProfileABC";
                repository.Save(profile);

                Assert.IsNotNull(repository.GetById("profileabc"));
                Assert.IsNull(repository.GetById("missing"));
            }
        }

        [TestMethod]
        public void Delete_RemovesExistingProfileAndRaisesChangedOnlyWhenRemoved()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                ConnectionProfileRepository repository = new ConnectionProfileRepository(paths);
                ConnectionProfile profile = CreateProfile("Primary", "sqlserver");
                repository.Save(profile);
                int changedCount = 0;
                string changedId = null;
                repository.Changed += (sender, args) =>
                {
                    changedCount++;
                    changedId = args.ProfileId;
                };

                repository.Delete("missing");
                repository.Delete(profile.Id);

                Assert.AreEqual(1, changedCount);
                Assert.AreEqual(profile.Id, changedId);
                Assert.AreEqual(0, new ConnectionProfileRepository(paths).GetAll().Count);
            }
        }

        [TestMethod]
        public void Save_ValidatesRequiredFieldsAndNormalizesNullParameters()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ConnectionProfileRepository repository = CreateRepository(temporary);
                Assert.ThrowsException<ArgumentNullException>(() => repository.Save(null));
                Assert.ThrowsException<ArgumentException>(() => repository.Save(CreateProfile(" ", "sqlserver")));
                Assert.ThrowsException<ArgumentException>(() => repository.Save(CreateProfile("Primary", " ")));
                ConnectionProfile profile = CreateProfile("Primary", "sqlserver");
                profile.Parameters = null;

                repository.Save(profile);

                Assert.AreEqual(0, repository.GetById(profile.Id).Parameters.Count);
            }
        }

        [TestMethod]
        public void Constructor_BacksUpCorruptFileAndStartsEmpty()
        {
            using (TemporaryDirectory temporary = new TemporaryDirectory())
            {
                ApplicationPaths paths = new ApplicationPaths(temporary.Path);
                paths.EnsureDirectories();
                File.WriteAllText(paths.ConnectionsFile, "not-json");

                ConnectionProfileRepository repository = new ConnectionProfileRepository(paths);

                Assert.AreEqual(0, repository.GetAll().Count);
                Assert.IsFalse(File.Exists(paths.ConnectionsFile));
                Assert.AreEqual(1, Directory.GetFiles(temporary.Path, "connections.json.corrupt-*").Length);
            }
        }

        private static ConnectionProfileRepository CreateRepository(TemporaryDirectory temporary)
        {
            return new ConnectionProfileRepository(new ApplicationPaths(temporary.Path));
        }

        private static ConnectionProfile CreateProfile(string name, string providerId)
        {
            return new ConnectionProfile
            {
                Name = name,
                ProviderId = providerId,
                Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["server"] = "localhost"
                }
            };
        }
    }
}
