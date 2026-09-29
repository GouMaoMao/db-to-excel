using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Contracts;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>在线程安全的内存集合与磁盘查询方案文件之间持久化查询方案。</summary>
    /// <remarks>
    /// 当前版本按“一方案一文件”保存到查询目录；首次启动会兼容读取旧版 queries.json 并迁移。
    /// 读取和保存均使用深复制，防止调用方绕过仓储修改内部状态。保存采用临时文件替换；损坏文件会重命名备份。
    /// </remarks>
    public sealed class QueryProfileRepository : IQueryProfileRepository
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private List<QueryProfile> _profiles;

        /// <summary>创建仓储并立即从磁盘加载查询方案。</summary>
        /// <param name="paths">查询文件路径来源。</param>
        public QueryProfileRepository(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _profiles = Load();
        }

        /// <inheritdoc/>
        public IReadOnlyList<QueryProfile> GetAll()
        {
            lock (_syncRoot)
            {
                return _profiles.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(Copy).ToList().AsReadOnly();
            }
        }

        /// <inheritdoc/>
        public QueryProfile GetById(string id)
        {
            lock (_syncRoot)
            {
                return Copy(_profiles.FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)));
            }
        }

        /// <inheritdoc/>
        public void Save(QueryProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.Id)) profile.Id = Guid.NewGuid().ToString("N");
            if (string.IsNullOrWhiteSpace(profile.Name)) throw new ArgumentException("查询方案名称不能为空。", nameof(profile));

            lock (_syncRoot)
            {
                QueryProfile copy = Copy(profile);
                copy.UpdatedUtc = DateTime.UtcNow;
                int index = _profiles.FindIndex(item =>
                    string.Equals(item.Id, copy.Id, StringComparison.OrdinalIgnoreCase));
                if (index >= 0) _profiles[index] = copy;
                else _profiles.Add(copy);
                SaveProfileUnsafe(copy);
            }
        }

        /// <inheritdoc/>
        public void Delete(string id)
        {
            lock (_syncRoot)
            {
                QueryProfile existing = _profiles.FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    _profiles.Remove(existing);
                    DeleteProfileUnsafe(existing.Id);
                }
            }
        }

        private List<QueryProfile> Load()
        {
            _paths.EnsureDirectories();

            List<QueryProfile> profiles = LoadFromDirectory();
            if (!File.Exists(_paths.QueriesFile)) return profiles;

            try
            {
                QueryDocument document = _serializer.Deserialize<QueryDocument>(
                    File.ReadAllText(_paths.QueriesFile, Encoding.UTF8));
                foreach (QueryProfile profile in document?.Profiles ?? new List<QueryProfile>())
                {
                    QueryProfile normalized = NormalizeLoadedProfile(profile);
                    int index = profiles.FindIndex(item => string.Equals(item.Id, normalized.Id, StringComparison.OrdinalIgnoreCase));
                    if (index >= 0) profiles[index] = normalized;
                    else profiles.Add(normalized);
                    SaveProfileUnsafe(normalized);
                }

                ArchiveMigratedLegacyFile();
                return profiles;
            }
            catch (Exception)
            {
                BackupCorrupt(_paths.QueriesFile);
                return profiles;
            }
        }

        private static QueryProfile Copy(QueryProfile source)
        {
            if (source == null) return null;
            return new QueryProfile
            {
                Id = source.Id,
                Name = source.Name,
                ProviderId = source.ProviderId,
                ConnectionProfileId = source.ConnectionProfileId,
                QueryText = source.QueryText,
                TargetSheetName = source.TargetSheetName,
                Description = source.Description,
                CreatedUtc = source.CreatedUtc,
                UpdatedUtc = source.UpdatedUtc,
                ProviderOptions = new Dictionary<string, string>(source.ProviderOptions ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase)
            };
        }

        private List<QueryProfile> LoadFromDirectory()
        {
            List<QueryProfile> profiles = new List<QueryProfile>();
            if (!Directory.Exists(_paths.QueryProfilesDirectory)) return profiles;

            foreach (string file in Directory.GetFiles(_paths.QueryProfilesDirectory, "*.json"))
            {
                try
                {
                    QueryProfileDocument document = _serializer.Deserialize<QueryProfileDocument>(File.ReadAllText(file, Encoding.UTF8));
                    QueryProfile profile = NormalizeLoadedProfile(document?.Profile);
                    if (profile != null) profiles.Add(profile);
                }
                catch (Exception)
                {
                    BackupCorrupt(file);
                }
            }

            return profiles;
        }

        private QueryProfile NormalizeLoadedProfile(QueryProfile profile)
        {
            if (profile == null) return null;

            QueryProfile copy = Copy(profile);
            if (string.IsNullOrWhiteSpace(copy.Id)) copy.Id = Guid.NewGuid().ToString("N");
            if (copy.ProviderOptions == null)
                copy.ProviderOptions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (copy.CreatedUtc == default(DateTime)) copy.CreatedUtc = DateTime.UtcNow;
            if (copy.UpdatedUtc == default(DateTime)) copy.UpdatedUtc = copy.CreatedUtc;
            return copy;
        }

        private void SaveProfileUnsafe(QueryProfile profile)
        {
            QueryProfile copy = Copy(profile);
            copy.ProviderOptions = copy.ProviderOptions ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string filePath = GetProfileFilePath(copy.Id);
            QueryProfileDocument document = new QueryProfileDocument
            {
                Version = 1,
                Profile = copy
            };
            AtomicWrite(filePath, _serializer.Serialize(document));
        }

        private void DeleteProfileUnsafe(string id)
        {
            string filePath = GetProfileFilePath(id);
            if (File.Exists(filePath)) File.Delete(filePath);
        }

        private string GetProfileFilePath(string profileId)
        {
            if (string.IsNullOrWhiteSpace(profileId))
                profileId = Guid.NewGuid().ToString("N");

            char[] invalid = Path.GetInvalidFileNameChars();
            string safe = new string(profileId.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray());
            return Path.Combine(_paths.QueryProfilesDirectory, safe + ".json");
        }

        private void AtomicWrite(string path, string content)
        {
            _paths.EnsureDirectories();
            string temporary = path + ".tmp";
            File.WriteAllText(temporary, content, Encoding.UTF8);
            if (File.Exists(path)) File.Replace(temporary, path, path + ".bak", true);
            else File.Move(temporary, path);
        }

        private void ArchiveMigratedLegacyFile()
        {
            string archive = _paths.QueriesFile + ".migrated-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            File.Move(_paths.QueriesFile, archive);
        }

        private static void BackupCorrupt(string path)
        {
            string backup = path + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            File.Move(path, backup);
        }

        /// <summary>定义单个查询方案文件的版本化对象。</summary>
        private sealed class QueryProfileDocument
        {
            public int Version { get; set; }
            public QueryProfile Profile { get; set; }
        }

        /// <summary>定义 queries.json 的版本化根对象。</summary>
        private sealed class QueryDocument
        {
            public int Version { get; set; }
            public List<QueryProfile> Profiles { get; set; }
        }
    }
}
