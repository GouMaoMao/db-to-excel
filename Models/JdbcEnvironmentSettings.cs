using System.Runtime.Serialization;

namespace DB2Sheet.Models
{
    /// <summary>保存本机 Java 可执行文件路径。</summary>
    /// <remarks>不属于通用设置页。用户在「JDBC 环境」中修改；厂商驱动 jar 保存在连接方案中。</remarks>
    [DataContract]
    public sealed class JdbcEnvironmentSettings
    {
        /// <summary>创建空的环境设置。</summary>
        public JdbcEnvironmentSettings()
        {
            JavaExecutable = string.Empty;
        }

        /// <summary>获取或设置 java.exe 的完整路径。为空时查询会尝试自动搜索本机。</summary>
        [DataMember(Order = 1)]
        public string JavaExecutable { get; set; }
    }
}
