namespace DB2Sheet.Providers
{
    /// <summary>
    /// 集中定义跨数据库提供程序共享的连接参数键名。
    /// </summary>
    /// <remarks>这些字符串会写入连接方案文件，修改会影响已有配置的兼容性。</remarks>
    public static class DatabaseParameterKeys
    {
        public const string Host = "host";
        public const string Port = "port";
        public const string Database = "database";
        public const string UserName = "userName";
        public const string Password = "password";
        public const string IntegratedSecurity = "integratedSecurity";
        public const string FilePath = "filePath";
        public const string Encrypt = "encrypt";
        public const string TrustServerCertificate = "trustServerCertificate";
    }
}
