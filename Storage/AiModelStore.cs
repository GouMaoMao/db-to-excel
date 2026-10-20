using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>把模型配置保存到 ai-models.json。密钥字段必须已经是密文。</summary>
    /// <remarks>读写都加锁。损坏文件会改名备份后当作没有配置。</remarks>
    public sealed class AiModelStore
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = CreateSerializer();
        private List<AiModelProfile> _profiles;

        /// <summary>创建存储并立即读取磁盘。</summary>
        /// <param name="paths">本地数据目录。</param>
        public AiModelStore(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _profiles = Load();
        }

        /// <summary>配置被保存或删除后触发。在锁外调用，便于界面刷新。</summary>
        public event EventHandler Changed;

        /// <summary>按显示名返回配置副本。</summary>
        /// <returns>当前全部配置。</returns>
        public IReadOnlyList<AiModelProfile> GetAll()
        {
            lock (_syncRoot)
            {
                return _profiles
                    .OrderBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                    .Select(Copy)
                    .ToList()
                    .AsReadOnly();
            }
        }

        /// <summary>保存一条配置。设为默认时会清掉其他配置的默认标记。</summary>
        /// <param name="profile">要保存的配置。标识为空时分配新标识。</param>
        public void Save(AiModelProfile profile)
        {
            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.DisplayName))
                throw new ArgumentException("显示名不能为空。", nameof(profile));
            if (string.IsNullOrWhiteSpace(profile.Endpoint))
                throw new ArgumentException("接口地址不能为空。", nameof(profile));
            bool hasName = !string.IsNullOrWhiteSpace(profile.ModelName);
            bool hasIds = AiModelCatalog.Ids(profile).Count > 0;
            if (!hasName && !hasIds)
                throw new ArgumentException("没有可用的型号。", nameof(profile));

            lock (_syncRoot)
            {
                AiModelProfile copy = Copy(profile);
                if (string.IsNullOrWhiteSpace(copy.Id))
                    copy.Id = Guid.NewGuid().ToString("N");
                if (copy.IsDefault)
                {
                    foreach (AiModelProfile item in _profiles)
                        item.IsDefault = false;
                }

                int index = _profiles.FindIndex(item =>
                    string.Equals(item.Id, copy.Id, StringComparison.OrdinalIgnoreCase));
                if (index >= 0) _profiles[index] = copy;
                else _profiles.Add(copy);
                SaveUnsafe();
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>删除一条配置。</summary>
        /// <param name="id">配置标识。</param>
        public void Delete(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return;
            lock (_syncRoot)
            {
                _profiles.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
                SaveUnsafe();
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        private List<AiModelProfile> Load()
        {
            try
            {
                if (!System.IO.File.Exists(_paths.AiModelsFile))
                    return new List<AiModelProfile>();
                AiModelDocument document = _serializer.Deserialize<AiModelDocument>(
                    System.IO.File.ReadAllText(_paths.AiModelsFile, System.Text.Encoding.UTF8));
                List<AiModelProfile> profiles = new List<AiModelProfile>();
                if (document?.Models == null) return profiles;
                foreach (AiModelProfile profile in document.Models)
                {
                    AiModelProfile copy = Normalize(profile);
                    if (copy != null) profiles.Add(copy);
                }

                return profiles;
            }
            catch (Exception)
            {
                AtomicFile.BackupCorrupt(_paths.AiModelsFile);
                return new List<AiModelProfile>();
            }
        }

        private static AiModelProfile Normalize(AiModelProfile profile)
        {
            if (profile == null || string.IsNullOrWhiteSpace(profile.DisplayName)) return null;
            AiModelProfile copy = Copy(profile);
            if (string.IsNullOrWhiteSpace(copy.Id))
                copy.Id = Guid.NewGuid().ToString("N");
            copy.Protocol = AiProtocolNames.ToStorage(AiProtocolNames.Parse(copy.Protocol));
            return copy;
        }

        private void SaveUnsafe()
        {
            AiModelDocument document = new AiModelDocument
            {
                Version = 1,
                Models = _profiles.Select(Copy).ToList()
            };
            AtomicFile.Write(_paths.AiModelsFile, _serializer.Serialize(document));
        }

        private static AiModelProfile Copy(AiModelProfile profile)
        {
            if (profile == null) return null;
            return new AiModelProfile
            {
                Id = profile.Id,
                DisplayName = profile.DisplayName,
                Protocol = profile.Protocol,
                Endpoint = profile.Endpoint,
                ModelName = profile.ModelName,
                ModelIds = profile.ModelIds == null ? null : new List<string>(profile.ModelIds),
                ProtectedKey = profile.ProtectedKey,
                IsDefault = profile.IsDefault
            };
        }

        private static JavaScriptSerializer CreateSerializer()
        {
            return new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
        }

        private sealed class AiModelDocument
        {
            public int Version { get; set; }
            public List<AiModelProfile> Models { get; set; }
        }
    }
}
