using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace DB2Sheet.Models
{
    /// <summary>以位标志描述数据源提供程序支持的功能。</summary>
    [Flags]
    public enum ProviderCapabilities
    {
        None = 0,
        Query = 1,
        Preview = 2,
        Streaming = 4,
        Cancellation = 8,
        ReadOnlySession = 16,
        ServerSideLimit = 32,
        Paging = 64
    }

    /// <summary>定义连接参数在编辑界面中的值类型和控件形式。</summary>
    public enum ParameterValueType
    {
        Text,
        Password,
        Integer,
        Boolean,
        FilePath,
        Choice
    }

    /// <summary>描述一个数据源连接参数的键名、显示方式、默认值和校验元数据。</summary>
    public sealed class ParameterDefinition
    {
        /// <summary>创建连接参数定义。</summary>
        /// <param name="key">用于持久化和读取的稳定键名。</param>
        /// <param name="displayName">显示给用户的名称。</param>
        /// <param name="valueType">参数值类型。</param>
        /// <param name="isRequired">是否必填。</param>
        /// <param name="defaultValue">默认值。</param>
        /// <param name="isSensitive">是否包含密码等敏感信息。</param>
        /// <param name="choices">选择类型允许的候选值。</param>
        /// <param name="visibleWhenKey">控制本参数是否显示的其他参数键；为空时始终显示。</param>
        /// <param name="visibleWhenValue">仅当控制参数等于该值时显示。</param>
        public ParameterDefinition(
            string key,
            string displayName,
            ParameterValueType valueType,
            bool isRequired = false,
            object defaultValue = null,
            bool isSensitive = false,
            IEnumerable<string> choices = null,
            string visibleWhenKey = null,
            string visibleWhenValue = null)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
            ValueType = valueType;
            IsRequired = isRequired;
            DefaultValue = defaultValue;
            IsSensitive = isSensitive;
            Choices = new List<string>(choices ?? new string[0]).AsReadOnly();
            VisibleWhenKey = visibleWhenKey ?? string.Empty;
            VisibleWhenValue = visibleWhenValue ?? string.Empty;
        }

        /// <summary>获取稳定键名。</summary>
        public string Key { get; }
        /// <summary>获取用户可见名称。</summary>
        public string DisplayName { get; }
        /// <summary>获取参数值类型。</summary>
        public ParameterValueType ValueType { get; }
        /// <summary>获取是否必填。</summary>
        public bool IsRequired { get; }
        /// <summary>获取默认值。</summary>
        public object DefaultValue { get; }
        /// <summary>获取是否属于不应记录日志的敏感值。</summary>
        public bool IsSensitive { get; }
        /// <summary>获取选择类型的候选值。</summary>
        public IReadOnlyList<string> Choices { get; }

        /// <summary>获取控制本参数显示的其他参数键。为空表示始终显示。</summary>
        public string VisibleWhenKey { get; }

        /// <summary>获取控制参数需匹配的值。</summary>
        public string VisibleWhenValue { get; }
    }

    /// <summary>描述查询结果中的一列及其 CLR 数据类型。</summary>
    /// <remarks>类型以程序集限定名称序列化；无法解析时回退为 <see cref="object"/>。</remarks>
    [DataContract]
    public sealed class ResultColumn
    {
        /// <summary>创建结果列定义。</summary>
        /// <param name="name">列名称。</param>
        /// <param name="dataType">列的 CLR 类型。</param>
        public ResultColumn(string name, Type dataType)
        {
            Name = name ?? string.Empty;
            DataTypeName = (dataType ?? typeof(object)).AssemblyQualifiedName;
        }

        [DataMember(Order = 1)]
        /// <summary>获取列名称。</summary>
        public string Name { get; private set; }

        [DataMember(Order = 2)]
        /// <summary>获取序列化后的程序集限定类型名。</summary>
        public string DataTypeName { get; private set; }

        /// <summary>解析并返回列的 CLR 类型。</summary>
        /// <returns>解析出的类型；失败时返回 <see cref="object"/>。</returns>
        public Type GetDataType()
        {
            return Type.GetType(DataTypeName, false) ?? typeof(object);
        }
    }

    /// <summary>表示从结果流读取的一批二维行列数据。</summary>
    public sealed class ResultBlock
    {
        /// <summary>创建结果块。</summary>
        /// <param name="values">以“行、列”索引的二维值数组。</param>
        /// <param name="rowCount">数组中有效的数据行数。</param>
        /// <param name="isCompleted">读取此块后结果流是否已结束。</param>
        public ResultBlock(object[,] values, int rowCount, bool isCompleted)
        {
            Values = values ?? new object[0, 0];
            RowCount = rowCount;
            IsCompleted = isCompleted;
        }

        /// <summary>获取二维数据数组。</summary>
        public object[,] Values { get; }
        /// <summary>获取有效数据行数。</summary>
        public int RowCount { get; }
        /// <summary>获取结果流是否已经读取完成。</summary>
        public bool IsCompleted { get; }
    }
}
