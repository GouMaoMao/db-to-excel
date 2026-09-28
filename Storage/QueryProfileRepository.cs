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
    /// <summary>在线程安全的内存集合与 queries.json 之间持久化查询方案。</summary>
    /// <remarks>
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
                SaveUnsafe();
            }
        }

        /// <inheritdoc/>
        public void Delete(string id)
        {
            lock (_syncRoot)
            {
                if (_profiles.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)) > 0)
                {
                    SaveUnsafe();
                }
            }
        }

        private List<QueryProfile> Load()
        {
            _paths.EnsureDirectories();
            if (!File.Exists(_paths.QueriesFile)) return new List<QueryProfile>();
            try
            {
                QueryDocument document = _serializer.Deserialize<QueryDocument>(
                    File.ReadAllText(_paths.QueriesFile, Encoding.UTF8));
                return (document?.Profiles ?? new List<QueryProfile>()).Select(Copy).ToList();
            }
            catch (Exception)
            {
                File.Move(_paths.QueriesFile, _paths.QueriesFile + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss"));
                return new List<QueryProfile>();
            }
        }

        private void SaveUnsafe()
        {
            string temporary = _paths.QueriesFile + ".tmp";
            string json = _serializer.Serialize(new QueryDocument
            {
                Version = 1,
                Profiles = _profiles.Select(Copy).ToList()
            });
            File.WriteAllText(temporary, json, Encoding.UTF8);
            if (File.Exists(_paths.QueriesFile)) File.Replace(temporary, _paths.QueriesFile, _paths.QueriesFile + ".bak", true);
            else File.Move(temporary, _paths.QueriesFile);
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

        /// <summary>定义 queries.json 的版本化根对象。</summary>
        private sealed class QueryDocument
        {
            public int Version { get; set; }
            public List<QueryProfile> Profiles { get; set; }
        }
    }
}
