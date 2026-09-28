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
    /// <summary>在线程安全的内存集合与 connections.json 之间持久化连接方案。</summary>
    /// <remarks>
    /// 对外始终返回深复制的不可变快照。保存采用临时文件替换并保留备份；损坏文件会重命名为 corrupt 备份后以空集合启动。
    /// </remarks>
    public sealed class ConnectionProfileRepository : IConnectionProfileRepository
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private List<ConnectionProfile> _profiles;

        /// <summary>创建仓储并立即从磁盘加载连接方案。</summary>
        /// <param name="paths">连接文件路径来源。</param>
        public ConnectionProfileRepository(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _profiles = Load();
        }

        /// <inheritdoc/>
        public event EventHandler<ConnectionProfilesChangedEventArgs> Changed;

        /// <inheritdoc/>
        public IReadOnlyList<ConnectionProfileSnapshot> GetAll()
        {
            lock (_syncRoot)
            {
                return _profiles
                    .OrderBy(profile => profile.Name, StringComparer.CurrentCultureIgnoreCase)
                    .Select(profile => profile.CreateSnapshot())
                    .ToList()
                    .AsReadOnly();
            }
        }

        /// <inheritdoc/>
        public ConnectionProfileSnapshot GetById(string id)
        {
            lock (_syncRoot)
            {
                ConnectionProfile profile = _profiles.FirstOrDefault(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                return profile?.CreateSnapshot();
            }
        }

        /// <inheritdoc/>
        public void Save(ConnectionProfile profile)
        {
            Validate(profile);
            string profileId;
            lock (_syncRoot)
            {
                ConnectionProfile copy = Copy(profile);
                profileId = copy.Id;
                int index = _profiles.FindIndex(item =>
                    string.Equals(item.Id, copy.Id, StringComparison.OrdinalIgnoreCase));
                if (index >= 0)
                {
                    _profiles[index] = copy;
                }
                else
                {
                    _profiles.Add(copy);
                }

                SaveUnsafe();
            }

            Changed?.Invoke(this, new ConnectionProfilesChangedEventArgs(profileId));
        }

        /// <inheritdoc/>
        public void Delete(string id)
        {
            bool removed;
            lock (_syncRoot)
            {
                removed = _profiles.RemoveAll(item =>
                    string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)) > 0;
                if (removed)
                {
                    SaveUnsafe();
                }
            }

            if (removed)
            {
                Changed?.Invoke(this, new ConnectionProfilesChangedEventArgs(id));
            }
        }

        private List<ConnectionProfile> Load()
        {
            _paths.EnsureDirectories();
            if (!File.Exists(_paths.ConnectionsFile))
            {
                return new List<ConnectionProfile>();
            }

            try
            {
                ConnectionDocument document = _serializer.Deserialize<ConnectionDocument>(
                    File.ReadAllText(_paths.ConnectionsFile, Encoding.UTF8));
                return (document?.Profiles ?? new List<ConnectionProfile>()).Select(Copy).ToList();
            }
            catch (Exception)
            {
                BackupCorrupt(_paths.ConnectionsFile);
                return new List<ConnectionProfile>();
            }
        }

        private void SaveUnsafe()
        {
            AtomicWrite(_paths.ConnectionsFile, _serializer.Serialize(new ConnectionDocument
            {
                Version = 1,
                Profiles = _profiles.Select(Copy).ToList()
            }));
        }

        private static void Validate(ConnectionProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.Id)) profile.Id = Guid.NewGuid().ToString("N");
            if (string.IsNullOrWhiteSpace(profile.Name)) throw new ArgumentException("连接方案名称不能为空。", nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.ProviderId)) throw new ArgumentException("必须选择数据源类型。", nameof(profile));
            if (profile.Parameters == null) profile.Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static ConnectionProfile Copy(ConnectionProfile source)
        {
            return new ConnectionProfile
            {
                Id = source.Id,
                Name = source.Name,
                ProviderId = source.ProviderId,
                Parameters = new Dictionary<string, string>(source.Parameters ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase)
            };
        }

        private void AtomicWrite(string path, string content)
        {
            _paths.EnsureDirectories();
            string temporary = path + ".tmp";
            File.WriteAllText(temporary, content, Encoding.UTF8);
            if (File.Exists(path)) File.Replace(temporary, path, path + ".bak", true);
            else File.Move(temporary, path);
        }

        private static void BackupCorrupt(string path)
        {
            string backup = path + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
            File.Move(path, backup);
        }

        /// <summary>定义 connections.json 的版本化根对象。</summary>
        private sealed class ConnectionDocument
        {
            public int Version { get; set; }
            public List<ConnectionProfile> Profiles { get; set; }
        }
    }
}
