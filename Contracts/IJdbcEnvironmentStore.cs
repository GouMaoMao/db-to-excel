using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>读取和保存本机 Java 路径。</summary>
    public interface IJdbcEnvironmentStore
    {
        /// <summary>读取环境设置。文件不存在时返回空路径。</summary>
        /// <returns>Java 路径设置。</returns>
        JdbcEnvironmentSettings Load();

        /// <summary>保存环境设置。</summary>
        /// <param name="settings">要保存的设置。</param>
        void Save(JdbcEnvironmentSettings settings);
    }
}
