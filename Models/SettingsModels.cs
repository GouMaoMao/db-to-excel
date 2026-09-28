using System;
using System.Globalization;

namespace DB2Sheet.Models
{
    /// <summary>定义设置值的归属范围。</summary>
    public enum SettingScope
    {
        User,
        Provider
    }

    /// <summary>非泛型设置定义基类，描述键名、默认值、解析和校验规则。</summary>
    public abstract class SettingDefinition
    {
        /// <summary>初始化设置定义的公共元数据。</summary>
        /// <param name="key">持久化使用的稳定键名。</param>
        /// <param name="displayName">用户可见名称。</param>
        /// <param name="category">设置界面分类。</param>
        /// <param name="scope">设置范围。</param>
        /// <param name="requiresRestart">更改后是否需要重启生效。</param>
        protected SettingDefinition(string key, string displayName, string category, SettingScope scope, bool requiresRestart)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
            Category = category ?? string.Empty;
            Scope = scope;
            RequiresRestart = requiresRestart;
        }

        /// <summary>获取稳定键名。</summary>
        public string Key { get; }
        /// <summary>获取用户可见名称。</summary>
        public string DisplayName { get; }
        /// <summary>获取设置分类。</summary>
        public string Category { get; }
        /// <summary>获取设置范围。</summary>
        public SettingScope Scope { get; }
        /// <summary>获取更改后是否需要重启。</summary>
        public bool RequiresRestart { get; }
        /// <summary>获取设置值的 CLR 类型。</summary>
        public abstract Type ValueType { get; }
        /// <summary>获取装箱后的默认值。</summary>
        public abstract object DefaultValue { get; }
        /// <summary>解析持久化字符串并同时验证值。</summary>
        /// <param name="value">持久化字符串。</param>
        /// <param name="result">解析后的装箱值。</param>
        /// <returns>解析且验证通过时为 <see langword="true"/>。</returns>
        public abstract bool TryParse(string value, out object result);
        /// <summary>检查运行时对象是否具有正确类型并满足业务约束。</summary>
        /// <param name="value">待验证对象。</param>
        /// <returns>有效时为 <see langword="true"/>。</returns>
        public abstract bool IsValid(object value);
        /// <summary>将有效设置值转换为区域性无关的持久化字符串。</summary>
        /// <param name="value">待序列化值。</param>
        /// <returns>持久化字符串。</returns>
        public abstract string Serialize(object value);
    }

    /// <summary>提供特定值类型的设置默认值、解析器和验证器。</summary>
    /// <typeparam name="T">设置值类型。</typeparam>
    public sealed class SettingDefinition<T> : SettingDefinition
    {
        private readonly Func<string, Tuple<bool, T>> _parser;
        private readonly Func<T, bool> _validator;

        /// <summary>创建强类型设置定义。</summary>
        /// <param name="key">稳定键名。</param>
        /// <param name="displayName">用户可见名称。</param>
        /// <param name="category">设置分类。</param>
        /// <param name="defaultValue">默认值。</param>
        /// <param name="parser">将字符串解析为成功标志和值的函数。</param>
        /// <param name="validator">可选的业务校验函数。</param>
        /// <param name="scope">设置范围。</param>
        /// <param name="requiresRestart">更改后是否需要重启。</param>
        public SettingDefinition(
            string key,
            string displayName,
            string category,
            T defaultValue,
            Func<string, Tuple<bool, T>> parser,
            Func<T, bool> validator = null,
            SettingScope scope = SettingScope.User,
            bool requiresRestart = false)
            : base(key, displayName, category, scope, requiresRestart)
        {
            Default = defaultValue;
            _parser = parser ?? throw new ArgumentNullException(nameof(parser));
            _validator = validator ?? (_ => true);
        }

        /// <summary>获取强类型默认值。</summary>
        public T Default { get; }
        /// <inheritdoc/>
        public override Type ValueType => typeof(T);
        /// <inheritdoc/>
        public override object DefaultValue => Default;

        /// <inheritdoc/>
        public override bool TryParse(string value, out object result)
        {
            Tuple<bool, T> parsed = _parser(value);
            result = parsed.Item2;
            return parsed.Item1 && _validator(parsed.Item2);
        }

        /// <inheritdoc/>
        public override bool IsValid(object value)
        {
            return value is T typed && _validator(typed);
        }

        /// <inheritdoc/>
        public override string Serialize(object value)
        {
            if (!(value is T typed))
            {
                throw new ArgumentException("设置值类型不匹配：" + Key, nameof(value));
            }

            return Convert.ToString(typed, CultureInfo.InvariantCulture);
        }
    }

    /// <summary>定义批量刷新任务采用串行还是并行查询。</summary>
    public enum BatchExecutionMode
    {
        Serial,
        Parallel
    }
}
