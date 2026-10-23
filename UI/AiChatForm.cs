using System;
using System.Drawing;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;

namespace DB2Sheet.UI
{
    /// <summary>SQL 执行失败时弹出的对话窗。与查询窗右侧共用同一份对话。</summary>
    /// <remarks>尺寸记在 session.formBounds 的 AiChat。关闭查询窗时由所有者关系一并关闭。</remarks>
    public sealed class AiChatForm : AppForm
    {
        private readonly AiChatControl _chat;

        /// <summary>创建弹出的对话窗。</summary>
        /// <param name="assistant">模型、Skill、对话和审计。</param>
        /// <param name="operationRunner">模型连接测试进度窗体运行器。</param>
        /// <param name="settings">记住窗体尺寸。</param>
        /// <param name="readEditor">发送时读取查询窗当前页签。</param>
        /// <param name="readMentions">输入 @ 时列出查询方案和当前预览。</param>
        public AiChatForm(
            AiAssistant assistant,
            IOperationRunner operationRunner,
            ISettingsStore settings,
            Func<AiEditorSnapshot> readEditor,
            Func<AiMentionSource> readMentions)
        {
            Text = AppPresentation.WindowTitle("AI 对话");
            Width = 440;
            Height = 640;
            MinimumSize = new Size(360, 480);
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            FormSizeMemory.Attach(this, settings, "AiChat");
            _chat = new AiChatControl(assistant, operationRunner, settings, readEditor, readMentions, null) { Dock = DockStyle.Fill };
            Controls.Add(_chat);
        }

        /// <summary>查询页签变化时刷新输入框上方的引用。</summary>
        public void NotifyEditorChanged()
        {
            _chat.NotifyEditorChanged();
        }

        /// <summary>新建一条标题为「SQL 报错」的对话，并把草稿填进输入框。</summary>
        /// <param name="draft">已经脱敏的 SQL 和错误摘要。不自动发送。</param>
        public void StartErrorConversation(string draft)
        {
            _chat.StartConversation("SQL 报错", draft);
        }
    }
}
