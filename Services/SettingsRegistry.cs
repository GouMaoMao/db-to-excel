using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using DB2Sheet.Contracts;
using DB2Sheet.Models;

namespace DB2Sheet.Services
{
    /// <summary>注册并按键名查找应用支持的设置定义。</summary>
    /// <remarks>
    /// 注册通常只发生在启动阶段；该类型本身不提供并发注册保护。
    /// <see cref="GetAll"/> 先按分类，再按 <see cref="SettingDefinition.SortOrder"/>，序号相同再按显示名。
    /// </remarks>
    public sealed class SettingsRegistry : ISettingsRegistry
    {
        private readonly Dictionary<string, SettingDefinition> _definitions =
            new Dictionary<string, SettingDefinition>(StringComparer.OrdinalIgnoreCase);

        /// <inheritdoc/>
        public IReadOnlyList<SettingDefinition> GetAll()
        {
            return _definitions.Values
                .OrderBy(definition => definition.Category)
                .ThenBy(definition => definition.SortOrder)
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
            registry.Register(CoreSettings.ResultSuccessColor);
            registry.Register(CoreSettings.ResultTruncatedColor);
            registry.Register(CoreSettings.ResultPendingColor);
            registry.Register(CoreSettings.ResultErrorColor);
            registry.Register(CoreSettings.BatchMode);
            registry.Register(CoreSettings.MaxParallelism);
            registry.Register(CoreSettings.QueryTimeoutSeconds);
            registry.Register(CoreSettings.ResultBlockSize);
            registry.Register(CoreSettings.LogLevel);
            registry.Register(CoreSettings.ActiveConnectionId);
            registry.Register(CoreSettings.WorkspaceSplitRatio);
            registry.Register(CoreSettings.RefreshSplitRatio);
            registry.Register(CoreSettings.FormBounds);
            registry.Register(CoreSettings.SqlSheetHeaderColor);
            registry.Register(CoreSettings.SqlSheetFormulaColor);
            registry.Register(CoreSettings.SqlSheetTabColor);
            registry.Register(CoreSettings.BeautifyAfterSave);
            registry.Register(CoreSettings.BeautifyAfterRefresh);
            registry.Register(CoreSettings.AiChatSplitRatio);
            registry.Register(CoreSettings.AiComposerInputHeight);
            registry.Register(CoreSettings.AiChatVisible);
            registry.Register(CoreSettings.AiModelChoice);
            registry.Register(CoreSettings.AiThinkingChoice);
            registry.Register(CoreSettings.AiSkillChoice);
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

        /// <summary>结果完整时的状态圆点和标题行底色。默认较浅，便于黑字辨认。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> ResultSuccessColor = ColorSetting(
            "excel.resultSuccessColor", "成功颜色", Color.FromArgb(101, 219, 91), 1, "Excel");

        /// <summary>结果被截断时的状态圆点和标题行底色。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> ResultTruncatedColor = ColorSetting(
            "excel.resultTruncatedColor", "截断颜色", Color.FromArgb(255, 184, 98), 2, "Excel");

        /// <summary>已勾选、尚未刷新的状态圆点颜色。不用于标题行。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> ResultPendingColor = ColorSetting(
            "excel.resultPendingColor", "待刷新颜色", Color.FromArgb(162, 162, 162), 1, "批量");

        /// <summary>刷新失败或参数、重名有误时的状态圆点颜色。不用于标题行。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> ResultErrorColor = ColorSetting(
            "excel.resultErrorColor", "失败颜色", Color.FromArgb(217, 57, 57), 2, "批量");

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

        /// <summary>
        /// 进度窗口日志列表的最低显示级别，取值与 <see cref="LogSeverity"/> 一致。
        /// 文件日志不读取此项，仍写入全部级别。持久化键名仍为 logging.level。
        /// </summary>
        public static readonly SettingDefinition<LogSeverity> LogLevel =
            new SettingDefinition<LogSeverity>(
                "logging.level", "进度日志级别", "日志",
                LogSeverity.Information,
                value => Tuple.Create(Enum.TryParse(value, true, out LogSeverity parsed), parsed));

        public static readonly SettingDefinition<string> ActiveConnectionId =
            new SettingDefinition<string>(
                "session.activeConnectionId", "当前连接", "会话", string.Empty,
                value => Tuple.Create(true, value ?? string.Empty));

        public static readonly SettingDefinition<int> WorkspaceSplitRatio = Integer(
            "session.workspaceSplitRatio", "左右分栏比例", "会话", 0, 0, 999);

        public static readonly SettingDefinition<int> RefreshSplitRatio = Integer(
            "session.refreshSplitRatio", "批量刷新分栏比例", "会话", 0, 0, 999);

        public static readonly SettingDefinition<string> FormBounds =
            new SettingDefinition<string>(
                "session.formBounds", "窗体尺寸", "会话", string.Empty,
                value => Tuple.Create(true, value ?? string.Empty));

        /// <summary>参数设置里 SQL 页格式这一页的分类名。批量刷新的「格式设置」也用它选中这一页。</summary>
        public const string SqlSheetFormatCategory = "sql页格式";

        /// <summary>第 1 行目标表名的颜色。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> SqlSheetHeaderColor = ColorSetting(
            "excel.sqlSheetHeaderColor", "表名颜色", SqlSheetPalette.Default.Header, 1);

        /// <summary>公式字体颜色。底色由其淡化，不单独保存。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> SqlSheetFormulaColor = ColorSetting(
            "excel.sqlSheetFormulaColor", "公式颜色", SqlSheetPalette.Default.Formula, 2);

        /// <summary>SQL 工作表底部标签的颜色。保存为 <c>R,G,B</c>。</summary>
        public static readonly SettingDefinition<Color> SqlSheetTabColor = ColorSetting(
            "excel.sqlSheetTabColor", "标签颜色", SqlSheetPalette.Default.Tab, 3);

        /// <summary>查询把 SQL 保存进 Excel 之后是否自动美化整张 SQL 页。默认开启。</summary>
        public static readonly SettingDefinition<bool> BeautifyAfterSave = Flag(
            "excel.beautifyAfterSave", "保存后美化", true, 4);

        /// <summary>批量刷新正常结束后是否自动美化整张 SQL 页。列多时较慢，默认关闭。</summary>
        public static readonly SettingDefinition<bool> BeautifyAfterRefresh = Flag(
            "batch.beautifyAfterRefresh", "刷新后美化", false, 5);

        /// <summary>SQL 查询窗右侧 AI 栏宽度占该栏所在区域的千分比。0 表示用户还没拖过，展开时对话栏为当时 SQL 列的 2/3。</summary>
        public static readonly SettingDefinition<int> AiChatSplitRatio = Integer(
            "session.aiChatSplitRatio", "AI 对话栏宽度比例", "会话", 0, 0, 999);

        /// <summary>AI 对话输入框高度，单位是像素。用户拖过才偏离默认。不出现在参数设置里。上限只拒绝明显无效的数字，界面上能拖到当前对话栏放得下为止。</summary>
        public static readonly SettingDefinition<int> AiComposerInputHeight = Integer(
            "session.aiComposerInputHeight", "AI 输入框高度", "会话", 72, 48, 16384);

        /// <summary>对话里上次选中的账号和子型号，形如「账号标识|型号」。不出现在参数设置里。</summary>
        public static readonly SettingDefinition<string> AiModelChoice =
            new SettingDefinition<string>(
                "session.aiModelChoice", "AI 型号选择", "会话", string.Empty,
                value => Tuple.Create(true, value ?? string.Empty));

        /// <summary>AI 对话里上次选中的思考强度，不出现在参数设置里。</summary>
        public static readonly SettingDefinition<string> AiThinkingChoice =
            new SettingDefinition<string>(
                "session.aiThinkingLevel", "AI 思考强度", "会话", AiThinkingLevels.Automatic,
                value =>
                {
                    string normalized = AiThinkingLevels.Normalize(value);
                    bool valid = string.IsNullOrWhiteSpace(value)
                        || string.Equals(value, AiThinkingLevels.Automatic, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(value, AiThinkingLevels.Low, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(value, AiThinkingLevels.Medium, StringComparison.OrdinalIgnoreCase)
                        || string.Equals(value, AiThinkingLevels.High, StringComparison.OrdinalIgnoreCase);
                    return Tuple.Create(valid, normalized);
                });

        /// <summary>对话里上次勾选的 Skill。多项用竖线分隔，一条都没勾时为空。不出现在参数设置里。</summary>
        public static readonly SettingDefinition<string> AiSkillChoice =
            new SettingDefinition<string>(
                "session.aiSkillChoice", "AI Skill 勾选", "会话", string.Empty,
                value => Tuple.Create(true, value ?? string.Empty));

        /// <summary>SQL 查询窗右侧 AI 栏是否展开。不出现在参数设置里。</summary>
        public static readonly SettingDefinition<bool> AiChatVisible =
            new SettingDefinition<bool>(
                "session.aiChatVisible", "AI 对话展开", "会话", false,
                value =>
                {
                    bool succeeded = bool.TryParse(value, out bool parsed);
                    return Tuple.Create(succeeded, parsed);
                });

        private static SettingDefinition<Color> ColorSetting(
            string key, string displayName, Color defaultColor, int sortOrder, string category = null)
        {
            return new SettingDefinition<Color>(
                key,
                displayName,
                category ?? SqlSheetFormatCategory,
                defaultColor,
                value =>
                {
                    if (!TryParseColor(value, out Color parsed))
                        return Tuple.Create(false, defaultColor);
                    return Tuple.Create(true, parsed);
                },
                sortOrder: sortOrder);
        }

        private static SettingDefinition<bool> Flag(string key, string displayName, bool defaultValue, int sortOrder)
        {
            return new SettingDefinition<bool>(
                key,
                displayName,
                SqlSheetFormatCategory,
                defaultValue,
                value =>
                {
                    if (!bool.TryParse(value, out bool parsed))
                        return Tuple.Create(false, defaultValue);
                    return Tuple.Create(true, parsed);
                },
                sortOrder: sortOrder);
        }

        private static bool TryParseColor(string text, out Color color)
        {
            color = Color.Empty;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string[] channels = text.Split(',');
            if (channels.Length != 3) return false;
            if (!int.TryParse(channels[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int red)
                || !int.TryParse(channels[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int green)
                || !int.TryParse(channels[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int blue))
            {
                return false;
            }

            if (red < 0 || red > 255 || green < 0 || green > 255 || blue < 0 || blue > 255) return false;
            color = Color.FromArgb(red, green, blue);
            return true;
        }

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
