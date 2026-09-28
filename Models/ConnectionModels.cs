using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Runtime.Serialization;

namespace DB2Sheet.Models
{
    /// <summary>表示可编辑、可序列化的数据库连接方案。</summary>
    /// <remarks>密码等敏感参数目前与普通参数一起保存，调用方不得将其写入日志。</remarks>
    [DataContract]
    public sealed class ConnectionProfile
    {
        /// <summary>创建具有新标识和空参数集合的连接方案。</summary>
        public ConnectionProfile()
        {
            Id = Guid.NewGuid().ToString("N");
            Name = string.Empty;
            ProviderId = string.Empty;
            Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        [DataMember(Order = 1)]
        /// <summary>获取或设置连接方案的稳定唯一标识。</summary>
        public string Id { get; set; }

        [DataMember(Order = 2)]
        /// <summary>获取或设置显示给用户的方案名称。</summary>
        public string Name { get; set; }

        [DataMember(Order = 3)]
        /// <summary>获取或设置数据源提供程序的内部标识。</summary>
        public string ProviderId { get; set; }

        [DataMember(Order = 4)]
        /// <summary>获取或设置以参数键名索引的连接参数。</summary>
        public Dictionary<string, string> Parameters { get; set; }

        /// <summary>复制当前值并创建不受后续编辑影响的只读快照。</summary>
        /// <returns>连接方案快照。</returns>
        public ConnectionProfileSnapshot CreateSnapshot()
        {
            return new ConnectionProfileSnapshot(Id, Name, ProviderId, Parameters);
        }
    }

    /// <summary>表示一次数据库操作使用的不可变连接配置快照。</summary>
    /// <remarks>构造时会复制参数字典，防止 UI 编辑与后台查询互相影响。</remarks>
    public sealed class ConnectionProfileSnapshot
    {
        /// <summary>从连接方案值创建快照。</summary>
        /// <param name="id">方案标识。</param>
        /// <param name="name">显示名称。</param>
        /// <param name="providerId">提供程序标识。</param>
        /// <param name="parameters">连接参数；键名不区分大小写。</param>
        public ConnectionProfileSnapshot(string id, string name, string providerId, IDictionary<string, string> parameters)
        {
            Id = id ?? string.Empty;
            Name = name ?? string.Empty;
            ProviderId = providerId ?? string.Empty;
            Parameters = new ReadOnlyDictionary<string, string>(
                new Dictionary<string, string>(parameters ?? new Dictionary<string, string>(), StringComparer.OrdinalIgnoreCase));
        }

        /// <summary>获取方案标识。</summary>
        public string Id { get; }
        /// <summary>获取显示名称。</summary>
        public string Name { get; }
        /// <summary>获取提供程序标识。</summary>
        public string ProviderId { get; }
        /// <summary>获取只读连接参数。</summary>
        public IReadOnlyDictionary<string, string> Parameters { get; }

        /// <summary>读取连接参数，不存在时返回指定默认值。</summary>
        /// <param name="key">参数键名。</param>
        /// <param name="defaultValue">参数不存在时返回的值。</param>
        /// <returns>参数值或默认值。</returns>
        public string GetValue(string key, string defaultValue = "")
        {
            return Parameters.TryGetValue(key, out string value) ? value : defaultValue;
        }
    }

    /// <summary>表示连接测试的成功状态、提示消息和耗时。</summary>
    public sealed class ConnectionTestResult
    {
        private ConnectionTestResult(bool succeeded, string message, TimeSpan elapsed)
        {
            Succeeded = succeeded;
            Message = message ?? string.Empty;
            Elapsed = elapsed;
        }

        /// <summary>获取测试是否成功。</summary>
        public bool Succeeded { get; }
        /// <summary>获取用于 UI 展示的结果消息。</summary>
        public string Message { get; }
        /// <summary>获取测试耗时。</summary>
        public TimeSpan Elapsed { get; }

        /// <summary>创建成功的连接测试结果。</summary>
        /// <param name="elapsed">测试耗时。</param>
        /// <returns>成功结果。</returns>
        public static ConnectionTestResult Success(TimeSpan elapsed)
        {
            return new ConnectionTestResult(true, "连接成功", elapsed);
        }

        /// <summary>创建失败的连接测试结果。</summary>
        /// <param name="message">失败原因。</param>
        /// <param name="elapsed">测试耗时。</param>
        /// <returns>失败结果。</returns>
        public static ConnectionTestResult Failure(string message, TimeSpan elapsed)
        {
            return new ConnectionTestResult(false, message, elapsed);
        }
    }

    /// <summary>为连接方案仓储的变更事件提供发生变化的方案标识。</summary>
    public sealed class ConnectionProfilesChangedEventArgs : EventArgs
    {
        /// <summary>创建连接方案变更事件参数。</summary>
        /// <param name="profileId">被保存或删除的方案标识。</param>
        public ConnectionProfilesChangedEventArgs(string profileId)
        {
            ProfileId = profileId;
        }

        /// <summary>获取发生变化的方案标识。</summary>
        public string ProfileId { get; }
    }
}
