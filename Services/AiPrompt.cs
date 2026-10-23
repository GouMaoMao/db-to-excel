using System;
using System.IO;
using System.Reflection;
using System.Text;

namespace DB2Sheet.Services
{
    /// <summary>从嵌入的 markdown 读出发给模型的说明，并填入当次的槽位。</summary>
    /// <remarks>
    /// 说明随程序集发布，不写到用户目录。用户能改的上下文仍然是 Skill 文件。
    /// 只做一次字符串替换，槽位的值里即使出现花括号也不会再被展开。
    /// </remarks>
    public static class AiPrompt
    {
        private const string Prefix = "DB2Sheet.Prompts.";
        private static readonly string SystemTemplate = Read("sql-assistant.system.md");
        private static readonly string UserTemplate = Read("sql-assistant.user.md");
        private static readonly string TabTemplate = Read("sql-assistant.tab.md");
        private static readonly string NoTabTemplate = Read("sql-assistant.no-tab.md");
        private static readonly string SkillTemplate = Read("sql-assistant.skill.md");
        private static readonly string ErrorTemplate = Read("sql-error.draft.md");
        private static readonly string ContextTemplate = Read("sql-assistant.context.md");
        private static readonly string TitleTemplate = Read("sql-assistant.title.md");

        /// <summary>产品自带的系统说明，不含 Skill。填入模型配置、子型号和所选思考强度。</summary>
        /// <param name="modelDisplayName">用户为当前模型配置填写的显示名。</param>
        /// <param name="modelId">对话里选中的子型号。模型按原文写入 SQL 注释，不是配置显示名。</param>
        /// <param name="thinkingLevel">当前 UI 选择的思考强度。</param>
        /// <returns>已替换模型配置、子型号和思考强度槽位的系统说明。</returns>
        public static string SystemText(string modelDisplayName, string modelId, string thinkingLevel)
        {
            return Fill(
                SystemTemplate,
                "modelDisplayName", modelDisplayName,
                "model", modelId,
                "thinkingLevel", thinkingLevel);
        }

        /// <summary>兼容只传子型号的调用，使用默认显示名和自动思考强度。</summary>
        /// <param name="modelId">对话里选中的子型号。</param>
        /// <returns>已替换提示词槽位的系统说明。</returns>
        public static string SystemText(string modelId)
        {
            return SystemText(string.Empty, modelId, "自动");
        }

        /// <summary>把 Skill 文档填上名称和正文。</summary>
        /// <param name="name">Skill 名称。</param>
        /// <param name="body">已经脱敏的 Skill 正文。</param>
        /// <returns>接在系统说明后面的一段。</returns>
        public static string Skill(string name, string body)
        {
            return Fill(SkillTemplate, "name", name, "body", body);
        }

        /// <summary>拼出这一次的用户消息。</summary>
        /// <param name="dialect">数据库类型说明。</param>
        /// <param name="hasTab">是否附带查询页签。</param>
        /// <param name="tabName">页签名称。</param>
        /// <param name="tabSql">已经脱敏的页签全文。</param>
        /// <param name="question">已经脱敏的用户输入。</param>
        /// <param name="context">已经拼好的额外引用。没有时为空。</param>
        /// <returns>发给模型的用户消息。</returns>
        public static string User(string dialect, bool hasTab, string tabName, string tabSql, string question, string context)
        {
            string tab = hasTab
                ? Fill(TabTemplate, "name", tabName, "sql", tabSql)
                : NoTabTemplate;
            return Fill(UserTemplate, "dialect", dialect, "tab", tab, "context", context, "question", question);
        }

        /// <summary>把一条引用填进上下文模板。</summary>
        /// <param name="label">芯片上的名称。</param>
        /// <param name="body">已经脱敏的正文。</param>
        /// <returns>一段引用。</returns>
        public static string Context(string label, string body)
        {
            return Fill(ContextTemplate, "label", label, "body", body);
        }

        /// <summary>填标题请求。只含当前标题和截断后的可见问答。</summary>
        /// <param name="title">当前标题。</param>
        /// <param name="dialog">近期可见问答。</param>
        /// <returns>单独发给模型的用户消息。</returns>
        public static string Title(string title, string dialog)
        {
            return Fill(TitleTemplate, "title", title, "dialog", dialog);
        }

        /// <summary>填报错草稿。调用方先完成脱敏和截断。</summary>
        /// <param name="sql">已经脱敏的 SQL。</param>
        /// <param name="error">已经脱敏并截断的错误摘要。</param>
        /// <returns>放进输入框的草稿。</returns>
        public static string ErrorDraft(string sql, string error)
        {
            return Fill(ErrorTemplate, "sql", sql, "error", error);
        }

        /// <summary>把模板里的 <c>{{名字}}</c> 换成对应的值。</summary>
        /// <param name="template">嵌入的说明。</param>
        /// <param name="slots">名字和值交替排列。</param>
        /// <returns>替换后的文本。</returns>
        public static string Fill(string template, params string[] slots)
        {
            string result = template ?? string.Empty;
            if (slots == null) return result;
            for (int index = 0; index + 1 < slots.Length; index += 2)
                result = result.Replace("{{" + (slots[index] ?? string.Empty) + "}}", slots[index + 1] ?? string.Empty);
            return result;
        }

        private static string Read(string fileName)
        {
            string name = Prefix + fileName;
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null)
                    throw new InvalidOperationException("缺少提示词 " + fileName + "。");
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    return reader.ReadToEnd().Trim();
            }
        }
    }
}
