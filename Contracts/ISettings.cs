using System;
using System.Collections.Generic;
using DB2Sheet.Models;

namespace DB2Sheet.Contracts
{
    /// <summary>保存应用支持的设置定义及其类型、默认值和校验规则。</summary>
    public interface ISettingsRegistry
    {
        /// <summary>获取全部设置定义。</summary>
        IReadOnlyList<SettingDefinition> GetAll();
        /// <summary>按稳定键名获取设置定义。</summary>
        /// <param name="key">设置键名。</param>
        SettingDefinition GetByKey(string key);
        /// <summary>注册或替换设置定义。</summary>
        /// <param name="definition">设置元数据和规则。</param>
        void Register(SettingDefinition definition);
    }

    /// <summary>读取、校验并持久化用户设置值。</summary>
    public interface ISettingsStore
    {
        /// <summary>一个或多个设置成功保存后触发。</summary>
        event EventHandler Changed;
        /// <summary>按非泛型定义读取设置值，缺失或无效时返回默认值。</summary>
        /// <param name="definition">设置定义。</param>
        object Get(SettingDefinition definition);
        /// <summary>按强类型定义读取设置值，缺失或无效时返回默认值。</summary>
        /// <typeparam name="T">设置值类型。</typeparam>
        /// <param name="definition">强类型设置定义。</param>
        T Get<T>(SettingDefinition<T> definition);
        /// <summary>校验并保存一个非泛型设置值。</summary>
        /// <param name="definition">设置定义。</param>
        /// <param name="value">待保存值。</param>
        void Set(SettingDefinition definition, object value);
        /// <summary>校验并保存一个强类型设置值。</summary>
        /// <typeparam name="T">设置值类型。</typeparam>
        /// <param name="definition">强类型设置定义。</param>
        /// <param name="value">待保存值。</param>
        void Set<T>(SettingDefinition<T> definition, T value);
        /// <summary>在一次持久化操作中保存多个设置。</summary>
        /// <param name="values">以设置键名索引的待保存值。</param>
        void SetMany(IReadOnlyDictionary<string, object> values);
        /// <summary>获取未经类型转换的持久化字符串快照。</summary>
        IReadOnlyDictionary<string, string> GetRawValues();
    }
}
