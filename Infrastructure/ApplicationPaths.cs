using System;
using System.IO;

namespace DB2Sheet.Infrastructure
{
    /// <summary>集中计算插件配置、查询、连接和日志文件的本地路径。</summary>
    /// <remarks>默认根目录位于当前用户的 LocalApplicationData 下；可注入自定义目录用于测试或隔离环境。</remarks>
    public sealed class ApplicationPaths
    {
        /// <summary>创建应用路径集合。</summary>
        /// <param name="rootDirectory">自定义根目录；为空时使用默认用户目录。</param>
        public ApplicationPaths(string rootDirectory = null)
        {
            RootDirectory = rootDirectory ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DB2Sheet");
        }

        /// <summary>获取所有应用数据的根目录。</summary>
        public string RootDirectory { get; }
        /// <summary>获取连接方案 JSON 文件路径。</summary>
        public string ConnectionsFile => Path.Combine(RootDirectory, "connections.json");
        /// <summary>获取查询方案 JSON 文件路径。</summary>
        public string QueriesFile => Path.Combine(RootDirectory, "queries.json");
        /// <summary>获取用户设置 JSON 文件路径。</summary>
        public string SettingsFile => Path.Combine(RootDirectory, "settings.json");
        /// <summary>获取日志目录路径。</summary>
        public string LogsDirectory => Path.Combine(RootDirectory, "Logs");

        /// <summary>确保根目录和日志目录存在。</summary>
        public void EnsureDirectories()
        {
            Directory.CreateDirectory(RootDirectory);
            Directory.CreateDirectory(LogsDirectory);
        }
    }
}
