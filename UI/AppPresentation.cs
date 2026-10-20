using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace DB2Sheet.UI
{
    /// <summary>集中定义所有用户可见品牌名称、版本展示和公共 UI 字体。</summary>
    /// <remarks>
    /// 更换插件展示名称时修改 <see cref="DisplayName"/> 即可。程序集名、持久化目录和 Ribbon 控件 ID 是独立的技术标识。
    /// 连接编辑窗的类型标题、类型介绍、框下说明、悬停和空框提示，以及参数设置窗的框下说明，也放在此处。
    /// </remarks>
    internal static class AppPresentation
    {
        public const string DisplayName = "db-to-excel";
        public const string IconResourceName = "DB2Sheet.Resources.rocket.ico";
        public const string DefaultFontName = "Microsoft YaHei UI";
        public const float DefaultFontSize = 9F;
        public const string LogFontName = "Consolas";
        public const float LogFontSize = 9F;
        public const string CodeFontName = "Microsoft YaHei UI Light";
        public const float CodeFontSize = 9.5F;

        /// <summary>SQL 工作表单元格使用的等宽字体。编辑器仍用 <see cref="CodeFontName"/>。</summary>
        public const string SqlSheetFontName = "Cascadia Code Light";

        /// <summary>SQL 工作表单元格的字号。</summary>
        public const float SqlSheetFontSize = 10F;

        /// <summary>对话栏灰色说明、Skill、引用和模型芯片的字号。比正文小一档。</summary>
        public const float AiComposerFontSize = 8.25F;

        /// <summary>连接编辑窗下拉的固定顺序。新建默认取第一项 SQL Server。</summary>
        public static readonly string[] ConnectionProviderOrder =
        {
            "sqlserver",
            "mysql",
            "postgresql",
            "sqlite",
            "duckdb",
            "jdbc"
        };

        /// <summary>显示名称空框提示。</summary>
        public const string ConnectionNameCue = "例如：ERP正式环境、数据中台XX空间";

        /// <summary>服务器地址空框提示。</summary>
        public const string ConnectionHostCue = "主机名或 IP";

        /// <summary>JDBC URL 下方说明的第一句。</summary>
        public const string JdbcUrlNote =
            "厂商规定的完整连接地址，写明连哪台机器、哪个库。向管理员索取后原样粘贴。";

        /// <summary>JDBC URL 下方的示例行。</summary>
        public const string JdbcUrlExample = "示例：jdbc:oracle:thin:@主机:1521:库名";

        /// <summary>驱动 jar 下方的说明。</summary>
        public const string DriverJarNote =
            "JDBC 驱动程序，数据库厂商提供的 .jar 文件。Java 加载它，才能按该库的协议建立连接。驱动含多个 jar 时全部添加。";

        /// <summary>驱动类下方的说明。</summary>
        public const string DriverClassNote =
            "该驱动程序中供 JDBC 加载的类名，例如 oracle.jdbc.OracleDriver。点「检测驱动类」从 jar 中读取；jar 未声明时按厂商文档填写。";

        /// <summary>JDBC 用户名下方的说明。</summary>
        public const string JdbcUserNameNote = "登录账号。有的库把账号写在 JDBC URL 里，这里可以留空。";

        /// <summary>jar 中读不到驱动类时的提示。</summary>
        public const string DriverClassMissingMessage = "jar 未声明驱动类，请按厂商文档填写。";

        /// <summary>未测试成功时保存按钮旁的说明。</summary>
        public const string SaveRequiresTestHint = "测试成功后才能保存";

        /// <summary>选中 JDBC 且未找到 Java 时的横幅文案。</summary>
        public const string JdbcJavaBanner =
            "通用连接（JDBC）需要本机 Java（JDK 8 或更高）。请先配置 JDBC 环境，再测试连接。";

        /// <summary>本地数据窗顶部说明。只查看路径，不改保存位置。</summary>
        public const string LocalDataIntro =
            "下面是当前用户本机上的配置和日志路径。这里只查看，不改保存位置。";

        /// <summary>本地数据窗「复制路径」的悬浮提示。</summary>
        public const string LocalDataCopyTip =
            "复制路径\r\n"
            + "把选中项的完整路径放进剪贴板。文件或目录还没生成时也可以复制。";

        /// <summary>本地数据窗「打开位置」的悬浮提示。</summary>
        public const string LocalDataOpenTip =
            "打开位置\r\n"
            + "在资源管理器中定位选中的文件，或打开选中的目录。还没生成时只说明，不新建。";

        /// <summary>本地数据窗「关闭」的悬浮提示。</summary>
        public const string LocalDataCloseTip =
            "关闭\r\n"
            + "关掉这个窗口。不改任何文件，也不改保存位置。";

        /// <summary>数据目录一行的悬浮提示。</summary>
        public const string LocalDataRootTip =
            "数据目录\r\n"
            + "当前 Windows 用户的本地数据根目录。下面各项都在这里。";

        /// <summary>连接方案一行的悬浮提示。</summary>
        public const string LocalDataConnectionsTip =
            "连接方案\r\n"
            + "全部连接方案写在这个文件里，含地址和登录信息。";

        /// <summary>查询方案一行的悬浮提示。</summary>
        public const string LocalDataQueriesTip =
            "查询方案\r\n"
            + "已保存的查询方案。一个方案一个文件。";

        /// <summary>参数设置一行的悬浮提示。会话状态也在同一文件，不单列。</summary>
        public const string LocalDataSettingsTip =
            "参数设置\r\n"
            + "参数设置里保存的项，以及不出现在设置界面的会话状态，都在这个文件里。";

        /// <summary>JDBC 环境一行的悬浮提示。</summary>
        public const string LocalDataJdbcTip =
            "JDBC 环境\r\n"
            + "本机 java.exe 的路径。厂商驱动 jar 写在连接方案里，不在这个文件。";

        /// <summary>模型配置一行的悬浮提示。</summary>
        public const string LocalDataModelsTip =
            "模型配置\r\n"
            + "模型的显示名、协议和接口地址。密钥是当前 Windows 用户的 DPAPI 密文。";

        /// <summary>Skill 目录一行的悬浮提示。</summary>
        public const string LocalDataSkillsTip =
            "Skill\r\n"
            + "一份 Skill 一个文件，只作为发给模型的文字。";

        /// <summary>对话目录一行的悬浮提示。</summary>
        public const string LocalDataConversationsTip =
            "对话\r\n"
            + "一条对话一个文件，含问答和查看请求详情用的原文。";

        /// <summary>日志目录一行的悬浮提示。日文件名写在详述里。</summary>
        public const string LocalDataLogsTip =
            "日志\r\n"
            + "按日写入 db2sheet-yyyyMMdd.log 和 ai-audit-yyyyMMdd.log。删除对话不删除审计日志。";

        /// <summary>对话页缓存一行的悬浮提示。</summary>
        public const string LocalDataWebViewTip =
            "对话页缓存\r\n"
            + "对话正文所用 WebView2 的浏览器数据，不在安装目录。";

        /// <summary>JDBC 转接程序一行的悬浮提示。首次使用 JDBC 才释放。</summary>
        public const string LocalDataBridgeTip =
            "JDBC 转接程序\r\n"
            + "嵌在插件里的转接 jar。第一次使用 JDBC 时才释放到数据目录。这不是连接方案中的驱动 jar。";

        /// <summary>模型设置列表末项「新增」的悬浮提示。</summary>
        public const string AiModelAddTip =
            "新增\r\n"
            + "清空右侧，开始填写一条新的模型配置。测试通过后才能保存。";

        /// <summary>模型设置列表右键「删除」的悬浮提示。</summary>
        public const string AiModelDeleteTip =
            "删除\r\n"
            + "删掉这条模型配置。已经保存的密钥一并移除。";

        /// <summary>模型设置里显示名下方的说明。</summary>
        public const string AiModelNameNote = "备注名称，方便记忆和管理。";

        /// <summary>模型设置里协议下方的说明。</summary>
        public const string AiModelProtocolNote = "多数选 OpenAI 兼容。";

        /// <summary>模型设置里接口地址下方的说明。</summary>
        public const string AiModelEndpointNote = "服务商提供。";

        /// <summary>模型设置里密钥下方的说明。</summary>
        public const string AiModelKeyNote = "在服务商的 API 平台生成。";

        /// <summary>Skill 设置列表末项「新增」的悬浮提示。</summary>
        public const string AiSkillAddTip =
            "新增\r\n"
            + "清空右侧，开始填写一条新的 Skill。";

        /// <summary>Skill 设置列表右键「删除」的悬浮提示。</summary>
        public const string AiSkillDeleteTip =
            "删除\r\n"
            + "删掉这条 Skill。";

        /// <summary>Skill 设置窗顶部的说明。正文只阅读，不执行。</summary>
        public const string AiSkillHelp =
            "Skill 是随问题一起发给模型的文字，用来交代口径、表含义或写法，让模型更懂需求。模型只阅读这段文字，不执行里面的 SQL，也不打开里面的链接。";

        /// <summary>模型设置窗顶部的协议说明。不做成可选模板。</summary>
        public const string AiProtocolHelp =
            "OpenAI 兼容：OpenAI、DeepSeek、MiniMax、Kimi、通义千问、智谱，以及大多数兼容网关。\r\n"
            + "Claude：Anthropic 官方接口上的 Claude。\r\n"
            + "Gemini：Google 官方 generateContent 接口上的 Gemini。\r\n"
            + "Claude 或 Gemini 若通过 OpenAI 兼容网关接入，协议仍选 OpenAI 兼容。";

        /// <summary>发送按钮的悬浮提示。第一行是概述，换行后写会带上什么、不会带上什么，以及输入框按键。</summary>
        public const string AiSendHint =
            "发送\r\n"
            + "将发送：你的问题、所选 Skill、输入框上方的查询页签全文、用 @ 附上的查询方案、文件或预览，以及数据库类型。删掉页签引用则不发送该页签 SQL。不自动发送查询结果、连接地址和密码。\r\n"
            + "Enter 发送，Shift+Enter 换行。输入 @ 可附加引用。";

        /// <summary>模型菜单末项。点开后弹出模型设置。</summary>
        public const string AiAddModelMenu = "新增其他模型";

        /// <summary>Skill 菜单末项。点开后弹出 Skill 设置。</summary>
        public const string AiAddSkillMenu = "新增其他 Skill";

        /// <summary>模型芯片的悬浮提示。第一行说明可以点开选择。</summary>
        public const string AiModelTip =
            "选择模型\r\n"
            + "这次提问用的子型号。名称末尾的下拉标记表示可以点开更换。点开可看到配置名称。菜单末项「新增其他模型」打开模型设置。到功能区「设置 → AI 设置 → 模型设置」也可以打开同一窗体。填写协议、接口地址和密钥，测试通过后保存。可用型号在测试时从接口读取。";

        /// <summary>Skill 芯片的悬浮提示。第一行说明可以点开选择。</summary>
        public const string AiSkillTip =
            "选择 Skill\r\n"
            + "和问题一起发给模型的说明，用来约定表、口径或写法。可以勾选多条。没选则不附带，芯片显示 Skills。只选一条时显示名称，多条时显示个数。菜单末项「新增其他 Skill」打开 Skill 设置。在功能区「设置 → AI 设置 → Skill 设置」里也可以编写。灰色图标表示这次不附带。";

        /// <summary>查询页签引用的悬浮提示。</summary>
        public const string AiSqlRefTip =
            "当前查询\r\n"
            + "发送时带上该页 SQL 全文。点末尾「×」后，这一次不发送该页 SQL。";

        /// <summary>用 @ 附上的预览芯片的悬浮提示。</summary>
        public const string AiPreviewChipTip =
            "预览引用\r\n"
            + "当前页的预览结果。点末尾「×」后，这一次不发送。";

        /// <summary>用户气泡下「查看请求详情」的文字。</summary>
        public const string AiSentLink = "查看请求详情";

        /// <summary>用户气泡下「查看请求详情」的悬浮提示。</summary>
        public const string AiSentLinkTip =
            "查看请求详情\r\n"
            + "打开这次组装给模型的原文。上方是请求地址、协议、模型和响应，接着是各段占比，点分类才展开正文。";

        /// <summary>请求详情里没有正文的分类显示的文字。</summary>
        public const string AiSentEmpty = "空";

        /// <summary>有输入 token 时，条形图下面的说明。{0} 是本次输入总数。</summary>
        public const string AiSentShareNote = "本次请求输入占用「{0}」tokens，各段 token 为估算。";

        /// <summary>还没有输入 token 时，条形图下面的说明。</summary>
        public const string AiSentEstimateNote = "本次请求的输入 token 尚未返回，各段 token 为估算。";

        /// <summary>请求详情「你的问题」标题。</summary>
        public const string AiSentQuestion = "你的问题";

        /// <summary>请求详情「默认说明」标题。</summary>
        public const string AiSentSystem = "默认说明";

        /// <summary>请求详情 Skill 分组标题。</summary>
        public const string AiSentSkill = "Skill";

        /// <summary>请求详情「数据库类型」标题。</summary>
        public const string AiSentDialect = "数据库类型";

        /// <summary>请求详情「当前查询页签」标题。</summary>
        public const string AiSentTab = "当前查询页签";

        /// <summary>请求详情「引用」分组标题。</summary>
        public const string AiSentCitation = "引用";

        /// <summary>请求详情「此前问答」标题。</summary>
        public const string AiSentHistory = "此前问答";

        /// <summary>此前问答下面的说明。再次发送时不带上一轮的页签、Skill 和引用。</summary>
        public const string AiSentHistoryNote =
            "再次发给模型的只有对话里看得见的问和答，不会把上一轮的页签 SQL、Skill 和引用再发一遍。";

        /// <summary>旧对话没有记下调用信息时，请求详情顶部的说明。</summary>
        public const string AiSentMissingCall = "这次没有记下请求地址和模型。";

        /// <summary>请求详情「请求地址」标题。</summary>
        public const string AiSentUrl = "请求地址";

        /// <summary>请求地址那一行。{0} 是实际 POST 的地址。</summary>
        public const string AiSentPostUrl = "POST {0}";

        /// <summary>请求详情「协议」标题。</summary>
        public const string AiSentProtocol = "协议";

        /// <summary>请求详情「模型」标题。</summary>
        public const string AiSentModel = "模型";

        /// <summary>请求详情「发送时间」标题。</summary>
        public const string AiSentTime = "发送时间";

        /// <summary>请求详情「响应」标题。</summary>
        public const string AiSentResponse = "响应";

        /// <summary>还没有 HTTP 状态时，响应那一行的文字。</summary>
        public const string AiSentPendingResponse = "尚未返回";

        /// <summary>收到回复后的响应。{0} 是状态码，{1} 是带千分位的毫秒数。</summary>
        public const string AiSentResponseValue = "HTTP {0}，耗时 {1} 毫秒";

        /// <summary>请求详情窗关闭按钮的悬浮提示。</summary>
        public const string AiSentCloseTip =
            "关闭\r\n"
            + "关掉这次请求详情。";

        /// <summary>代码块复制成功后，显示在复制图标旁。</summary>
        public const string AiCodeCopied = "已复制";

        /// <summary>代码块未能写入剪贴板时，显示在复制图标旁。</summary>
        public const string AiCodeCopyFailed = "复制失败";

        /// <summary>用 @ 附上的方案或文件芯片的悬浮提示。</summary>
        public const string AiAttachmentChipTip =
            "引用\r\n"
            + "这次会发给模型的内容。点末尾「×」后去掉。";

        /// <summary>输入框下方轮换的灰色说明。多行文本框不能显示空框提示，所以单独放一行。</summary>
        /// <remarks>每次创建对话栏抽一条，本次打开期间不换。有多条时不连续抽到同一条。</remarks>
        public static readonly string[] AiComposerHints =
        {
            "输入 @ 可附加查询方案、文件或当前预览。",
            "复制AI生成的SQL代码到SQL编辑器区域执行。使用者必须理解SQL代码逻辑，再使用。",
            "模型切换：不同模型在生成质量、响应速度、Token消耗上存在区别。",
            "Skills：为AI注入场景知识，显著提升生成SQL的准确性。",
            "Token：对话文本的计费单位，描述越简洁消耗越低；不建议引用MB级别的文件给AI，会大幅增加token消耗。"
        };

        private static readonly Random HintRandom = new Random();
        private static int _lastAiComposerHint = -1;

        /// <summary>抽一条与上次不同的输入框说明。</summary>
        /// <remarks>只在创建对话栏的界面线程上调用。对话栏按返回的这一句显示到关掉为止。</remarks>
        /// <returns>这次对话栏固定显示的那一句。</returns>
        public static string PickAiComposerHint()
        {
            int count = AiComposerHints.Length;
            if (count <= 0) return string.Empty;
            if (count == 1 || _lastAiComposerHint < 0)
            {
                int first = count == 1 ? 0 : HintRandom.Next(count);
                _lastAiComposerHint = first;
                return AiComposerHints[first];
            }

            int index = HintRandom.Next(count - 1);
            if (index >= _lastAiComposerHint)
                index++;
            _lastAiComposerHint = index;
            return AiComposerHints[index];
        }

        /// <summary>查询标题栏预览按钮的悬浮提示。</summary>
        public const string PreviewQueryTip =
            "预览\r\n"
            + "执行当前页签的 SQL，把结果放在下方预览区。行数受「最大预览行数」限制。";

        /// <summary>查询标题栏写入按钮的悬浮提示。</summary>
        public const string ExportQueryTip =
            "写入 Sheet\r\n"
            + "把当前查询结果写入 Excel。写入前确认目标表名。";

        /// <summary>预览表右键「复制」的悬浮提示。</summary>
        public const string PreviewCopyTip =
            "复制\r\n"
            + "把选中的格子放进剪贴板。没有选区时复制全部数据行。不含行号和标题。";

        /// <summary>预览表右键「复制带标题」的悬浮提示。</summary>
        public const string PreviewCopyHeadersTip =
            "复制带标题\r\n"
            + "在数据前加上两行：字段编码和类型短名。有选区时只含选中的列和行。";

        /// <summary>预览表右键「全选」的悬浮提示。</summary>
        public const string PreviewSelectAllTip =
            "全选\r\n"
            + "选中预览表里的全部数据行。";

        /// <summary>预览表右键「导出为 Excel」的悬浮提示。</summary>
        public const string PreviewExportTip =
            "导出为 Excel\r\n"
            + "把当前已经查出的预览存成 xlsx。先选择保存路径。不重新查询，也不改当前工作簿。";

        /// <summary>预览列标题右键「复制标题」的悬浮提示。</summary>
        public const string PreviewCopyTitleTip =
            "复制标题\r\n"
            + "复制这一列的字段编码和类型短名。";

        /// <summary>查询标题栏保存按钮的悬浮提示。</summary>
        public const string SaveQueryTip =
            "保存方案\r\n"
            + "把当前页签的 SQL 写入查询方案。未保存的新建页签在这时才成为方案。";

        /// <summary>查询标题栏 AI 按钮的悬浮提示。</summary>
        public const string ToggleAiTip =
            "AI\r\n"
            + "展开或收起右侧对话栏。";

        /// <summary>对话页签条上新建按钮的悬浮提示。</summary>
        public const string AiNewConversationTip =
            "新建对话\r\n"
            + "开始一条空白对话，与右键「新建」相同。";

        /// <summary>对话页签条上历史按钮的悬浮提示。</summary>
        public const string AiArchivedTip =
            "已归档\r\n"
            + "列出已收起的对话。选中一条后把它放回页签。";

        /// <summary>对话栏右上角关闭按钮的悬浮提示。</summary>
        public const string AiCloseChatTip =
            "收起\r\n"
            + "收起右侧对话栏，与再点一次「AI」相同。";

        /// <summary>没有对话时，编辑区中央按钮上的文字。</summary>
        public const string EmptyConversationCaption = "请新建对话";

        /// <summary>没有对话时，中央按钮的悬浮提示。</summary>
        public const string EmptyConversationTip =
            "请新建对话\r\n"
            + "开始一条空白对话。";

        /// <summary>没有已归档对话时，历史菜单里的不可用说明。</summary>
        public const string AiNoArchivedConversation = "没有已归档的对话";

        /// <summary>查询页签条上新建按钮的悬浮提示。</summary>
        public const string NewQueryTip =
            "新建查询\r\n"
            + "打开一个空白页签，与菜单「新建查询」相同。保存之后才写入查询方案。";

        /// <summary>没有查询页签时，编辑区中央按钮上的文字。</summary>
        public const string EmptyQueryCaption = "请新建查询";

        /// <summary>没有查询页签时，中央按钮的悬浮提示。</summary>
        public const string EmptyQueryTip =
            "请新建查询\r\n"
            + "打开一个空白页签。保存之后才成为查询方案。";

        /// <summary>一个连接方案都没有时的说明。用于阻止打开查询或批量刷新，以及执行时列表已经为空。</summary>
        public const string NoConnectionPrompt =
            "还没有数据库连接。请先点击功能区「连接管理」，再点「新建」。";

        /// <summary>已有连接但未设为当前连接时，预览、写入或批量刷新前的说明。</summary>
        public const string SetCurrentConnectionPrompt =
            "请在左侧「数据库连接」列表中选中一个连接，右键选择「设置为当前连接」。";

        private const int EmSetCueBanner = 0x1501;

        /// <summary>读取当前程序集的产品版本，供 Ribbon 等界面展示。</summary>
        /// <remarks>优先使用三段的 AssemblyFileVersion，不读取程序集文件路径，避免 VSTO 缓存路径读不到版本。特性缺失时回退程序集版本号。</remarks>
        public static string FileVersion
        {
            get
            {
                Assembly assembly = typeof(AppPresentation).Assembly;
                AssemblyFileVersionAttribute attribute = Attribute.GetCustomAttribute(
                    assembly,
                    typeof(AssemblyFileVersionAttribute)) as AssemblyFileVersionAttribute;
                if (attribute != null && !string.IsNullOrWhiteSpace(attribute.Version))
                    return attribute.Version;

                return assembly.GetName().Version.ToString();
            }
        }

        /// <summary>为 SQL 编辑器创建中英文同一字面高度的常规代码字体；优先细体雅黑，缺失时回退。</summary>
        /// <returns>已解析到本机已安装字体族的字体实例。调用方负责释放。</returns>
        public static Font CreateCodeFont()
        {
            string[] candidates = { "Microsoft YaHei UI Light", "DengXian Light", "Microsoft YaHei UI" };
            foreach (string name in candidates)
            {
                using (Font probe = new Font(name, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point))
                {
                    if (!string.Equals(probe.FontFamily.Name, name, StringComparison.OrdinalIgnoreCase))
                        continue;
                }

                return new Font(name, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point);
            }

            return new Font(DefaultFontName, CodeFontSize, FontStyle.Regular, GraphicsUnit.Point);
        }

        /// <summary>生成统一格式的功能窗体标题。</summary>
        /// <param name="featureName">功能名称；为空时只返回产品展示名称。</param>
        /// <returns>“产品名 - 功能名”格式的标题。</returns>
        public static string WindowTitle(string featureName)
        {
            return string.IsNullOrWhiteSpace(featureName)
                ? DisplayName
                : DisplayName + " - " + featureName;
        }

        /// <summary>连接类型在界面上的标题。JDBC 显示为「通用连接（JDBC）」。</summary>
        /// <param name="providerId">Provider 标识。</param>
        /// <returns>用户可见标题；未知标识时回退原值。</returns>
        public static string ConnectionProviderTitle(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId)) return string.Empty;
            switch (providerId.Trim().ToLowerInvariant())
            {
                case "sqlserver": return "SQL Server";
                case "mysql": return "MySQL";
                case "postgresql": return "PostgreSQL";
                case "sqlite": return "SQLite";
                case "duckdb": return "DuckDB";
                case "jdbc": return "通用连接（JDBC）";
                default: return providerId;
            }
        }

        /// <summary>选中某连接类型时显示在下拉下方的一句介绍，只说这种库的特点。</summary>
        /// <param name="providerId">Provider 标识。</param>
        /// <returns>一句介绍；未知标识为空。</returns>
        public static string ConnectionProviderHint(string providerId)
        {
            if (string.IsNullOrWhiteSpace(providerId)) return string.Empty;
            switch (providerId.Trim().ToLowerInvariant())
            {
                case "sqlserver":
                    return "微软的企业数据库，常见于公司业务系统，可用当前 Windows 账号登录。";
                case "mysql":
                    return "使用最广的开源数据库，多见于网站和一般业务系统。";
                case "postgresql":
                    return "开源数据库，擅长复杂查询和较严格的数据约束。";
                case "sqlite":
                    return "不需要数据库服务，整个库是本机上的一个文件。";
                case "duckdb":
                    return "本机上的一个文件，专为分析查询设计，适合统计和汇总。";
                case "jdbc":
                    return "通用连接方式，用来连接这里没有单独列出的数据库，例如 Oracle，以及 Hive 等数据仓库。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>输入框下方的灰色说明。没有说明的字段返回空，窗体不另占一行。</summary>
        /// <param name="providerId">当前数据源类型。用户名说明只在 JDBC 下出现。</param>
        /// <param name="fieldKey">字段键，如 jdbcUrl、driverJars、integratedSecurity。</param>
        /// <returns>框下说明；不需要说明时为空。</returns>
        public static string ConnectionFieldNote(string providerId, string fieldKey)
        {
            if (string.IsNullOrWhiteSpace(fieldKey)) return string.Empty;
            bool jdbc = string.Equals(providerId, "jdbc", StringComparison.OrdinalIgnoreCase);
            switch (fieldKey.Trim())
            {
                case "jdbcUrl":
                    return JdbcUrlNote + Environment.NewLine + JdbcUrlExample;
                case "driverJars":
                    return DriverJarNote;
                case "driverClass":
                    return DriverClassNote;
                case "userName":
                    return jdbc ? JdbcUserNameNote : string.Empty;
                case "integratedSecurity":
                    return "勾选后用当前 Windows 账号登录，用户名和密码可留空。";
                case "encrypt":
                    return "传输是否加密。没有管理员要求时保持默认。";
                case "trustServerCertificate":
                    return "内网或自签证书连不上时再勾选。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>连接编辑字段的悬停说明。已有框下说明的字段与说明使用同一句话。</summary>
        /// <param name="providerId">当前数据源类型，用于区分 JDBC 用户名和其他库的用户名。</param>
        /// <param name="fieldKey">字段键，如 host、jdbcUrl，或 name、provider、test、save。</param>
        /// <returns>悬停说明；未知键为空。</returns>
        public static string ConnectionFieldTip(string providerId, string fieldKey)
        {
            string note = ConnectionFieldNote(providerId, fieldKey);
            if (!string.IsNullOrEmpty(note)) return note;
            if (string.IsNullOrWhiteSpace(fieldKey)) return string.Empty;
            switch (fieldKey.Trim())
            {
                case "name":
                    return "在连接列表里显示的名字，方便认出环境。";
                case "provider":
                    return "要连接的数据库种类；常见库选上面几项，没有再选其他（JDBC）。";
                case "host":
                    return "数据库所在机器的主机名或 IP。";
                case "port":
                    return "一般用默认即可，除非管理员另有要求。";
                case "database":
                    return "要查询的库名。";
                case "userName":
                    return "登录该库的账号；勾选 Windows 集成认证时可留空。";
                case "password":
                    return "登录该库的密码；勾选 Windows 集成认证时可留空。";
                case "filePath":
                    return "本机数据库文件路径，可点浏览选择。";
                case "test":
                    return "用当前填写内容试连；通过后才能保存。";
                case "save":
                    return "测试通过后才能保存；保存后可在查询里使用。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>参数设置窗里某一项下方的灰色说明。会话状态不出现在设置窗，这里也不写。</summary>
        /// <param name="key">设置键，与 <c>CoreSettings</c> 中的持久化键相同。</param>
        /// <returns>一句定义；未知键为空，窗体不另占一行。</returns>
        public static string SettingNote(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return string.Empty;
            switch (key.Trim())
            {
                case "query.maxPreviewRows":
                    return "SQL 查询窗预览时最多读取并显示的行数。";
                case "excel.maxExportRows":
                    return "写入 Excel 的结果数据行上限，不含标题行。超出的行被截断。";
                case "excel.resultSuccessColor":
                    return "批量刷新的成功状态，以及结果完整时标题行的底色。标题为黑字，宜用较浅的颜色。";
                case "excel.resultTruncatedColor":
                    return "批量刷新的截断状态，以及结果达到最大导出行数被截断时标题行的底色。";
                case "excel.resultPendingColor":
                    return "已勾选、尚未刷新的状态圆点。";
                case "excel.resultErrorColor":
                    return "刷新失败，或参数、重名有误时的状态圆点。";
                case "batch.mode":
                    return "Serial 为串行，一项查询结束后再查下一项。Parallel 为并行，多项同时查询。";
                case "batch.maxParallelism":
                    return "并行时同时查询的任务个数。串行时不起作用。";
                case "query.timeoutSeconds":
                    return "单条查询允许运行的最长时间，超时后该查询失败。";
                case "query.resultBlockSize":
                    return "从数据库读取结果时每次取回的行数。只影响读取节奏，不改变最终行数。";
                case "logging.level":
                    return "进度窗口只显示不低于所选级别的日志。Debug 最详细，Error 只显示错误。写入文件的日志仍保留全部级别。";
                case "excel.sqlSheetHeaderColor":
                    return "第 1 行目标表名的字体颜色，并加粗显示。// 之后的参数保持灰色。";
                case "excel.sqlSheetCommentColor":
                    return "以 -- 或 /* */ 书写的注释文本。";
                case "excel.sqlSheetFormulaColor":
                    return "Excel 公式单元格的字体颜色。底色按该颜色向白色淡化。";
                case "excel.sqlSheetKeywordColor":
                    return "SELECT、FROM、WHERE 等 SQL 关键字。";
                case "excel.sqlSheetLiteralColor":
                    return "单引号字符串的字体颜色，包括两侧引号。双引号、反引号和方括号中的标识符保持黑色。";
                case "excel.sqlSheetTabColor":
                    return "SQL 工作表标签的颜色。";
                case "excel.beautifyAfterSave":
                    return "SQL 查询将语句保存到 Excel 后，自动美化整张 SQL 页。";
                case "batch.beautifyAfterRefresh":
                    return "批量刷新正常结束后，自动美化整张 SQL 页。列数较多时耗时较长，默认关闭。";
                default:
                    return string.Empty;
            }
        }

        /// <summary>给文本框设置灰色空框提示。句柄尚未创建时等创建后再设。</summary>
        /// <param name="textBox">目标文本框。</param>
        /// <param name="cue">提示文字。</param>
        public static void SetCueBanner(TextBox textBox, string cue)
        {
            if (textBox == null || string.IsNullOrEmpty(cue)) return;
            if (textBox.IsHandleCreated)
            {
                SendMessage(textBox.Handle, EmSetCueBanner, (IntPtr)1, cue);
                return;
            }

            EventHandler onCreated = null;
            onCreated = (sender, args) =>
            {
                textBox.HandleCreated -= onCreated;
                if (!textBox.IsDisposed && textBox.IsHandleCreated)
                {
                    SendMessage(textBox.Handle, EmSetCueBanner, (IntPtr)1, cue);
                }
            };
            textBox.HandleCreated += onCreated;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
    }
}
