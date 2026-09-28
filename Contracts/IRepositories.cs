using System;
using System.Collections.Generic;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>管理可持久化的数据库连接方案。</summary>
    public interface IConnectionProfileRepository
    {
        /// <summary>连接方案保存或删除后触发。</summary>
        event EventHandler<ConnectionProfilesChangedEventArgs> Changed;
        /// <summary>获取全部连接方案的不可变快照。</summary>
        IReadOnlyList<ConnectionProfileSnapshot> GetAll();
        /// <summary>按标识获取连接方案快照，不存在时返回 <see langword="null"/>。</summary>
        /// <param name="id">连接方案标识。</param>
        ConnectionProfileSnapshot GetById(string id);
        /// <summary>新增或更新连接方案。</summary>
        /// <param name="profile">要持久化的可变连接模型。</param>
        void Save(ConnectionProfile profile);
        /// <summary>删除指定连接方案。</summary>
        /// <param name="id">连接方案标识。</param>
        void Delete(string id);
    }

    /// <summary>管理可复用的 SQL 查询方案。</summary>
    public interface IQueryProfileRepository
    {
        /// <summary>获取全部查询方案。</summary>
        IReadOnlyList<QueryProfile> GetAll();
        /// <summary>按标识获取查询方案，不存在时返回 <see langword="null"/>。</summary>
        /// <param name="id">查询方案标识。</param>
        QueryProfile GetById(string id);
        /// <summary>新增或更新查询方案。</summary>
        /// <param name="profile">要持久化的查询方案。</param>
        void Save(QueryProfile profile);
        /// <summary>删除指定查询方案。</summary>
        /// <param name="id">查询方案标识。</param>
        void Delete(string id);
    }
}
