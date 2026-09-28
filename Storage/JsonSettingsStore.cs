using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Contracts;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>依据设置注册表校验类型，并将值以字符串形式保存到 settings.json。</summary>
    /// <remarks>
    /// 所有内存访问和文件写入均通过锁串行化。无效或缺失值读取为定义中的默认值；损坏文件会重命名备份后回退为空设置。
    /// </remarks>
    public sealed class JsonSettingsStore : ISettingsStore
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly ISettingsRegistry _registry;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private Dictionary<string, string> _values;

        /// <summary>创建设置存储并立即读取磁盘文件。</summary>
        /// <param name="paths">设置文件路径来源。</param>
        /// <param name="registry">用于类型解析和约束校验的设置注册表。</param>
        public JsonSettingsStore(ApplicationPaths paths, ISettingsRegistry registry)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
            _values = Load();
        }

        /// <inheritdoc/>
        public event EventHandler Changed;

        /// <inheritdoc/>
        public object Get(SettingDefinition definition)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            lock (_syncRoot)
            {
                if (_values.TryGetValue(definition.Key, out string text) &&
                    definition.TryParse(text, out object parsed))
                {
                    return parsed;
                }

                return definition.DefaultValue;
            }
        }

        /// <inheritdoc/>
        public T Get<T>(SettingDefinition<T> definition)
        {
            return (T)Get((SettingDefinition)definition);
        }

        /// <inheritdoc/>
        public void Set(SettingDefinition definition, object value)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            if (!definition.IsValid(value))
            {
                throw new ArgumentOutOfRangeException(nameof(value), "设置值不符合约束：" + definition.Key);
            }

            lock (_syncRoot)
            {
                _values[definition.Key] = definition.Serialize(value);
                SaveUnsafe();
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        /// <inheritdoc/>
        public void Set<T>(SettingDefinition<T> definition, T value)
        {
            Set((SettingDefinition)definition, value);
        }

        /// <inheritdoc/>
        public void SetMany(IReadOnlyDictionary<string, object> values)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            Dictionary<string, string> serialized = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object> pair in values)
            {
                SettingDefinition definition = _registry.GetByKey(pair.Key) ??
                    throw new ArgumentException("未注册的设置项：" + pair.Key, nameof(values));
                if (!definition.IsValid(pair.Value))
                {
                    throw new ArgumentOutOfRangeException(nameof(values), "设置值不符合约束：" + definition.Key);
                }
                serialized[definition.Key] = definition.Serialize(pair.Value);
            }

            lock (_syncRoot)
            {
                foreach (KeyValuePair<string, string> pair in serialized)
                {
                    _values[pair.Key] = pair.Value;
                }
                SaveUnsafe();
            }

            Changed?.Invoke(this, EventArgs.Empty);
        }

        /// <inheritdoc/>
        public IReadOnlyDictionary<string, string> GetRawValues()
        {
            lock (_syncRoot)
            {
                return new Dictionary<string, string>(_values, StringComparer.OrdinalIgnoreCase);
            }
        }

        private Dictionary<string, string> Load()
        {
            _paths.EnsureDirectories();
            if (!File.Exists(_paths.SettingsFile))
            {
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            try
            {
                string json = File.ReadAllText(_paths.SettingsFile, Encoding.UTF8);
                SettingsDocument document = _serializer.Deserialize<SettingsDocument>(json);
                return new Dictionary<string, string>(
                    document?.Values ?? new Dictionary<string, string>(),
                    StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                string backup = _paths.SettingsFile + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                File.Move(_paths.SettingsFile, backup);
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        private void SaveUnsafe()
        {
            _paths.EnsureDirectories();
            SettingsDocument document = new SettingsDocument
            {
                Version = 1,
                Values = new Dictionary<string, string>(_values, StringComparer.OrdinalIgnoreCase)
            };

            string temporaryFile = _paths.SettingsFile + ".tmp";
            File.WriteAllText(temporaryFile, _serializer.Serialize(document), Encoding.UTF8);
            if (File.Exists(_paths.SettingsFile))
            {
                File.Replace(temporaryFile, _paths.SettingsFile, _paths.SettingsFile + ".bak", true);
            }
            else
            {
                File.Move(temporaryFile, _paths.SettingsFile);
            }
        }

        /// <summary>定义 settings.json 的版本化根对象。</summary>
        private sealed class SettingsDocument
        {
            public int Version { get; set; }
            public Dictionary<string, string> Values { get; set; }
        }
    }
}
