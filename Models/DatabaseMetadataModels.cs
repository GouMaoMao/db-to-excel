using System;

namespace DB2Sheet.Models
{
    /// <summary>区分数据库目录中的表和视图对象。</summary>
    public enum DatabaseObjectKind
    {
        /// <summary>可直接查询的持久表。</summary>
        Table,
        /// <summary>由查询定义的数据库视图。</summary>
        View
    }

    /// <summary>描述当前账号可见的一个数据库。</summary>
    public sealed class DatabaseMetadata
    {
        /// <summary>创建数据库元数据。</summary>
        /// <param name="name">数据库名称。</param>
        public DatabaseMetadata(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("数据库名称不能为空。", nameof(name));
            Name = name;
        }

        /// <summary>获取数据库名称。</summary>
        public string Name { get; }
    }

    /// <summary>描述数据库内一个带架构归属的表或视图。</summary>
    public sealed class DatabaseObjectMetadata
    {
        /// <summary>创建数据库对象元数据。</summary>
        /// <param name="schemaName">对象所属架构；数据库不支持架构时使用空字符串。</param>
        /// <param name="name">对象名称。</param>
        /// <param name="kind">对象类型。</param>
        /// <param name="comment">数据库中的对象注释。没有注释时使用空字符串。</param>
        public DatabaseObjectMetadata(string schemaName, string name, DatabaseObjectKind kind, string comment = null)
        {
            if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("数据库对象名称不能为空。", nameof(name));
            SchemaName = schemaName ?? string.Empty;
            Name = name;
            Kind = kind;
            Comment = comment ?? string.Empty;
        }

        /// <summary>获取对象所属架构。</summary>
        public string SchemaName { get; }
        /// <summary>获取对象名称。</summary>
        public string Name { get; }
        /// <summary>获取对象类型。</summary>
        public DatabaseObjectKind Kind { get; }
        /// <summary>获取数据库对象注释。没有注释时为空字符串。</summary>
        public string Comment { get; }
    }
}
