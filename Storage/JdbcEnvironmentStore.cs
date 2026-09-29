using System;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using DB2Sheet.Contracts;
using DB2Sheet.Infrastructure;
using DB2Sheet.Models;

namespace DB2Sheet.Storage
{
    /// <summary>把 JDBC 环境设置保存到 jdbc-environment.json。</summary>
    /// <remarks>
    /// 读写通过锁串行化。文件损坏时会重命名备份并回退到空路径。
    /// 保存使用临时文件替换。旧字段如 JavaMode、DriverDirectory 会被忽略。
    /// </remarks>
    public sealed class JdbcEnvironmentStore : IJdbcEnvironmentStore
    {
        private readonly object _syncRoot = new object();
        private readonly ApplicationPaths _paths;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();

        /// <summary>创建存储。</summary>
        /// <param name="paths">应用数据目录。</param>
        public JdbcEnvironmentStore(ApplicationPaths paths)
        {
            _paths = paths ?? throw new ArgumentNullException(nameof(paths));
        }

        /// <summary>读取环境设置。文件不存在时返回空路径。</summary>
        /// <returns>当前设置。</returns>
        public JdbcEnvironmentSettings Load()
        {
            lock (_syncRoot)
            {
                JdbcEnvironmentSettings settings = ReadUnsafe();
                settings.JavaExecutable = settings.JavaExecutable ?? string.Empty;
                return settings;
            }
        }

        /// <summary>保存环境设置。</summary>
        /// <param name="settings">要保存的 Java 路径。</param>
        public void Save(JdbcEnvironmentSettings settings)
        {
            if (settings == null) throw new ArgumentNullException(nameof(settings));
            lock (_syncRoot)
            {
                JdbcEnvironmentDocument document = new JdbcEnvironmentDocument
                {
                    Version = 3,
                    JavaExecutable = settings.JavaExecutable ?? string.Empty
                };
                _paths.EnsureDirectories();
                string temporaryFile = _paths.JdbcEnvironmentFile + ".tmp";
                File.WriteAllText(temporaryFile, _serializer.Serialize(document), Encoding.UTF8);
                if (File.Exists(_paths.JdbcEnvironmentFile))
                {
                    File.Replace(temporaryFile, _paths.JdbcEnvironmentFile, _paths.JdbcEnvironmentFile + ".bak", true);
                }
                else
                {
                    File.Move(temporaryFile, _paths.JdbcEnvironmentFile);
                }
            }
        }

        private JdbcEnvironmentSettings ReadUnsafe()
        {
            if (!File.Exists(_paths.JdbcEnvironmentFile))
            {
                return new JdbcEnvironmentSettings();
            }

            try
            {
                string json = File.ReadAllText(_paths.JdbcEnvironmentFile, Encoding.UTF8);
                JdbcEnvironmentDocument document = _serializer.Deserialize<JdbcEnvironmentDocument>(json);
                return new JdbcEnvironmentSettings
                {
                    JavaExecutable = document?.JavaExecutable ?? string.Empty
                };
            }
            catch (Exception)
            {
                string backup = _paths.JdbcEnvironmentFile + ".corrupt-" + DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                File.Move(_paths.JdbcEnvironmentFile, backup);
                return new JdbcEnvironmentSettings();
            }
        }

        private sealed class JdbcEnvironmentDocument
        {
            public int Version { get; set; }
            public string JavaExecutable { get; set; }
        }
    }
}
