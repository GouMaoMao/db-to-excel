using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>注册并按键名查找应用支持的设置定义。</summary>
    /// <remarks>注册通常只发生在启动阶段；该类型本身不提供并发注册保护。</remarks>
    public sealed class SettingsRegistry : ISettingsRegistry
    {
        private readonly Dictionary<string, SettingDefinition> _definitions =
            new Dictionary<string, SettingDefinition>(StringComparer.OrdinalIgnoreCase);

        /// <inheritdoc/>
        public IReadOnlyList<SettingDefinition> GetAll()
        {
            return _definitions.Values
                .OrderBy(definition => definition.Category)
                .ThenBy(definition => definition.DisplayName)
                .ToList()
                .AsReadOnly();
        }

        /// <inheritdoc/>
        public SettingDefinition GetByKey(string key)
        {
            _definitions.TryGetValue(key ?? string.Empty, out SettingDefinition definition);
            return definition;
        }

        /// <inheritdoc/>
        public void Register(SettingDefinition definition)
        {
            if (definition == null)
            {
                throw new ArgumentNullException(nameof(definition));
            }

            if (_definitions.ContainsKey(definition.Key))
            {
                throw new InvalidOperationException("设置项已注册：" + definition.Key);
            }

            _definitions.Add(definition.Key, definition);
        }

        /// <summary>创建并注册插件当前支持的全部默认设置。</summary>
        /// <returns>可直接传给设置存储和设置窗体的注册表。</returns>
        public static SettingsRegistry CreateDefault()
        {
            SettingsRegistry registry = new SettingsRegistry();
            registry.Register(CoreSettings.MaxPreviewRows);
            registry.Register(CoreSettings.MaxExportRows);
            registry.Register(CoreSettings.BatchMode);
            registry.Register(CoreSettings.MaxParallelism);
            registry.Register(CoreSettings.QueryTimeoutSeconds);
            registry.Register(CoreSettings.ResultBlockSize);
            registry.Register(CoreSettings.LogLevel);
            return registry;
        }
    }

    /// <summary>集中定义核心功能使用的强类型设置项和约束。</summary>
    /// <remarks>键名会持久化到设置文件，修改时需要考虑旧版本兼容性。</remarks>
    public static class CoreSettings
    {
        private const int ExcelMaximumDataRows = 1048575;

        public static readonly SettingDefinition<int> MaxPreviewRows = Integer(
            "query.maxPreviewRows", "最大预览行数", "查询", 100, 1, 100000);

        public static readonly SettingDefinition<int> MaxExportRows = Integer(
            "excel.maxExportRows", "最大导出数据行数", "Excel", ExcelMaximumDataRows, 1, ExcelMaximumDataRows);

        public static readonly SettingDefinition<BatchExecutionMode> BatchMode =
            new SettingDefinition<BatchExecutionMode>(
                "batch.mode", "批量执行模式", "批量",
                BatchExecutionMode.Serial,
                value => Tuple.Create(Enum.TryParse(value, true, out BatchExecutionMode parsed), parsed));

        public static readonly SettingDefinition<int> MaxParallelism = Integer(
            "batch.maxParallelism", "最大并发数", "批量", 4, 1, 32);

        public static readonly SettingDefinition<int> QueryTimeoutSeconds = Integer(
            "query.timeoutSeconds", "查询超时（秒）", "查询", 300, 1, 86400);

        public static readonly SettingDefinition<int> ResultBlockSize = Integer(
            "query.resultBlockSize", "结果分块行数", "查询", 2000, 100, 20000);

        public static readonly SettingDefinition<string> LogLevel =
            new SettingDefinition<string>(
                "logging.level", "日志级别", "日志", "Information",
                value => Tuple.Create(!string.IsNullOrWhiteSpace(value), value),
                value => value == "Debug" || value == "Information" || value == "Warning" || value == "Error");

        private static SettingDefinition<int> Integer(
            string key, string displayName, string category, int defaultValue, int minimum, int maximum)
        {
            return new SettingDefinition<int>(
                key,
                displayName,
                category,
                defaultValue,
                value =>
                {
                    bool succeeded = int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed);
                    return Tuple.Create(succeeded, parsed);
                },
                value => value >= minimum && value <= maximum);
        }
    }
}
