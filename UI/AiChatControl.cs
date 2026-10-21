using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DB2Sheet.Contracts;
using DB2Sheet.Models;
using DB2Sheet.Services;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace DB2Sheet.UI
{
    /// <summary>查询窗右侧和报错弹窗共用的对话控件。</summary>
    /// <remarks>
    /// 正文由 WebView2 显示。发送在界面线程发起，等待网络时不占用界面线程，完成后回到界面线程更新页面和页签。
    /// 两个控件可以同时打开，各自记住当前对话；列表变化通过存储事件刷新。
    /// </remarks>
    public sealed class AiChatControl : UserControl
    {
        private const int MaxFileBytes = 64 * 1024;
        private const int ConversationTabIconDefault = 0;
        private const int ConversationTabIconPending = 1;
        private const int ConversationTabIconError = 2;
        private const int ConversationTabIconUnread = 3;

        private readonly AiAssistant _assistant;
        private readonly ISettingsStore _settings;
        private readonly Func<AiEditorSnapshot> _readEditor;
        private readonly Func<AiMentionSource> _readMentions;
        private readonly Action _closePanel;
        private readonly TabControl _tabs = CloseMark.CreateTabs();
        private readonly ImageList _conversationTabIcons = CreateConversationTabIcons();
        private readonly Panel _empty = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = SystemColors.Window,
            Visible = false
        };
        private readonly Button _emptyButton;
        private readonly WebView2 _view;
        private readonly Label _browserFallback = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = SystemColors.GrayText,
            Visible = false
        };
        private readonly Button _modelButton;
        private readonly Button _skillButton;
        private readonly Button _sqlRef;
        private readonly FlowLayoutPanel _chipRow;
        private readonly Panel _composer;
        private readonly Dictionary<string, Image> _mentionIcons = new Dictionary<string, Image>(StringComparer.OrdinalIgnoreCase);
        private readonly List<Attachment> _attachments = new List<Attachment>();
        private readonly MentionPopup _mention;
        private readonly ToolTip _tips = new ToolTip();
        private readonly TextBox _input = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            BorderStyle = BorderStyle.FixedSingle,
            AcceptsReturn = true,
            BackColor = SystemColors.Window
        };
        private readonly Button _send = new Button { Text = "发送", AutoSize = true, UseVisualStyleBackColor = true };
        private readonly Font _chipFont = new Font(AppPresentation.DefaultFontName, AppPresentation.AiComposerFontSize, FontStyle.Regular, GraphicsUnit.Point);
        private readonly Image _sendIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_send.png", 24);
        private readonly Image _sendIconGray = DatabaseTreeImageCatalog.LoadActionIcon("icon_send.png", 24, true);
        private readonly Label _mentionHint;
        private Panel _modelRow;
        private readonly ContextMenuStrip _tabMenu = new ContextMenuStrip();
        private readonly ToolStripMenuItem _renameItem;
        private readonly ToolStripMenuItem _archiveItem;
        private readonly ToolStripMenuItem _deleteItem;
        private readonly ToolStripMenuItem _archivedMenu;
        private readonly ContextMenuStrip _modelMenu = new ContextMenuStrip();
        private readonly ContextMenuStrip _skillMenu = new ContextMenuStrip();
        private readonly ContextMenuStrip _historyMenu = new ContextMenuStrip();
        private readonly List<Image> _tabIcons = new List<Image>();
        private readonly HashSet<string> _errorConversationIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _unreadConversationIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Image _addIcon = DatabaseTreeImageCatalog.LoadActionIcon("icon_add.png");
        private readonly List<ModelPick> _models = new List<ModelPick>();
        private readonly List<SkillPick> _skills = new List<SkillPick>();
        private bool _binding;
        private bool _suppressMention;
        private bool _sending;
        private int _mentionStart = -1;
        private bool _omitSql;
        private bool _browserStarted;
        private bool _browserReady;
        private bool _browserFailed;
        private string _currentId;
        private string _modelId;
        private string _modelName;
        private readonly List<string> _skillIds = new List<string>();
        private bool _skillChoiceLoaded;
        private bool _keepSkillMenu;
        private string _pendingConversationId;
        private string _pendingUser;
        private List<AiSentSection> _pendingSent;
        private AiSentCall _pendingCall;
        private int _titleSerial;

        /// <summary>创建对话页签图标集合。顺序固定为默认、等待、失败和未读。</summary>
        /// <returns>调用方负责释放的图像列表。</returns>
        private static ImageList CreateConversationTabIcons()
        {
            ImageList images = new ImageList
            {
                ColorDepth = ColorDepth.Depth32Bit,
                ImageSize = new Size(16, 16),
                TransparentColor = Color.Transparent
            };
            try
            {
                images.Images.Add(DatabaseTreeImageCatalog.LoadActionIcon("icon_talk.png", 16));
                images.Images.Add(DatabaseTreeImageCatalog.LoadActionIcon("icon_ing.png", 16));
                images.Images.Add(DatabaseTreeImageCatalog.LoadActionIcon("icon_error.png", 16));
                images.Images.Add(DatabaseTreeImageCatalog.LoadActionIcon("icon_unread.png", 16));
                return images;
            }
            catch
            {
                images.Dispose();
                throw;
            }
        }

        /// <summary>创建对话控件。</summary>
        /// <param name="assistant">模型、Skill、对话和审计。</param>
        /// <param name="settings">记住上次选中的子型号。</param>
        /// <param name="readEditor">点击发送时读取当前页签。由查询窗提供。</param>
        /// <param name="readMentions">输入 @ 时列出查询方案和当前预览。</param>
        /// <param name="closePanel">右上角关闭图标要做的事。弹出的对话窗不传，那里用窗体自己的关闭。</param>
        public AiChatControl(
            AiAssistant assistant,
            ISettingsStore settings,
            Func<AiEditorSnapshot> readEditor,
            Func<AiMentionSource> readMentions,
            Action closePanel)
        {
            _assistant = assistant ?? throw new ArgumentNullException(nameof(assistant));
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
            _readEditor = readEditor ?? throw new ArgumentNullException(nameof(readEditor));
            _readMentions = readMentions ?? throw new ArgumentNullException(nameof(readMentions));
            _closePanel = closePanel;
            AiChatBrowserHost.PrepareLoader();
            _view = new WebView2 { Dock = DockStyle.Fill };
            Font = new Font(AppPresentation.DefaultFontName, AppPresentation.DefaultFontSize, FontStyle.Regular, GraphicsUnit.Point);
            Dock = DockStyle.Fill;
            BackColor = SystemColors.Window;

            _modelButton = QuietButton();
            _skillButton = QuietButton();
            _sqlRef = ReferenceChip(MentionKind.Query);
            _sqlRef.Margin = new Padding(4, 2, 4, 0);
            _modelButton.TextImageRelation = TextImageRelation.ImageBeforeText;
            _modelButton.Padding = new Padding(2, 0, 10, 0);
            _skillButton.TextImageRelation = TextImageRelation.ImageBeforeText;
            _modelButton.Click += (sender, args) => _modelMenu.Show(_modelButton, new Point(0, 0), ToolStripDropDownDirection.AboveRight);
            _skillButton.Click += (sender, args) => _skillMenu.Show(_skillButton, new Point(0, _skillButton.Height));
            _skillMenu.Closing += SkillMenuClosing;
            _sqlRef.MouseDown += SqlRefMouseDown;
            _input.KeyDown += InputKeyDown;
            _input.TextChanged += InputTextChanged;
            _send.Click += SendButtonClick;
            _mention = new MentionPopup(AcceptMention, MentionIcon);
            _emptyButton = HintButton.Create(AppPresentation.EmptyConversationCaption);
            _emptyButton.Click += (sender, args) => StartConversation("新对话", string.Empty, false);
            _emptyButton.MouseUp += EmptyMouseUp;
            _tips.SetToolTip(_emptyButton, AppPresentation.EmptyConversationTip);
            _empty.Controls.Add(_emptyButton);
            _empty.Resize += (sender, args) => CenterEmptyButton();
            _tips.SetToolTip(_modelButton, AppPresentation.AiModelTip);
            _tips.SetToolTip(_skillButton, AppPresentation.AiSkillTip);
            _tips.SetToolTip(_sqlRef, AppPresentation.AiSqlRefTip);
            _tips.SetToolTip(_send, AppPresentation.AiSendHint);

            _renameItem = new ToolStripMenuItem("重命名", null, (sender, args) => Guard(RenameCurrent));
            _archiveItem = new ToolStripMenuItem("归档", null, (sender, args) => Guard(ArchiveCurrent));
            _deleteItem = new ToolStripMenuItem("删除", null, (sender, args) => Guard(DeleteCurrent));
            _archivedMenu = new ToolStripMenuItem("已归档");
            _tabMenu.Items.Add("新建", null, (sender, args) => StartConversation("新对话", string.Empty, false));
            _tabMenu.Items.Add(_renameItem);
            _tabMenu.Items.Add(_archiveItem);
            _tabMenu.Items.Add(_deleteItem);
            _tabMenu.Items.Add(_archivedMenu);
            _tabs.MouseDown += TabsMouseDown;
            _tabs.MouseUp += TabsMouseUp;
            _tabs.SelectedIndexChanged += TabsSelectedIndexChanged;
            _empty.MouseUp += EmptyMouseUp;

            _chipRow = ChoiceRow();
            _chipRow.WrapContents = true;
            _chipRow.AutoSize = false;
            _chipRow.Controls.Add(_skillButton);
            _chipRow.Controls.Add(_sqlRef);
            Panel inputHost = new Panel { Dock = DockStyle.Top, Height = 72, Padding = new Padding(8, 0, 8, 0), BackColor = SystemColors.Control };
            inputHost.Controls.Add(_input);
            _mentionHint = new Label
            {
                Dock = DockStyle.Top,
                Height = 20,
                Text = AppPresentation.PickAiComposerHint(),
                Font = _chipFont,
                ForeColor = SystemColors.GrayText,
                TextAlign = ContentAlignment.TopLeft,
                Padding = new Padding(8, 2, 8, 0),
                BackColor = SystemColors.Control
            };
            _send.Font = Font;
            _send.Image = _sendIcon;
            _send.TextImageRelation = TextImageRelation.ImageBeforeText;
            _send.ImageAlign = ContentAlignment.MiddleLeft;
            _send.Padding = new Padding(10, 4, 14, 4);
            _send.Margin = Padding.Empty;
            Size sendSize = _send.GetPreferredSize(Size.Empty);
            _send.AutoSize = false;
            _send.Width = SendButtonWidth();
            _send.Dock = DockStyle.Right;
            _modelButton.Location = new Point(8, 4);
            _modelRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = Math.Max(36, sendSize.Height),
                Padding = new Padding(0, 0, 8, 0),
                BackColor = SystemColors.Control
            };
            _modelRow.Controls.Add(_modelButton);
            _modelRow.Controls.Add(_send);
            _composer = new Panel { Dock = DockStyle.Bottom, Height = 156, BackColor = SystemColors.Control };
            _composer.Controls.Add(_modelRow);
            _composer.Controls.Add(_mentionHint);
            _composer.Controls.Add(inputHost);
            _composer.Controls.Add(_chipRow);
            _composer.Resize += (sender, args) => LayoutComposer();
            Panel line = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = SystemColors.ControlDark };
            Panel transcript = new Panel { Dock = DockStyle.Fill, BackColor = SystemColors.Window };
            transcript.Controls.Add(_view);
            transcript.Controls.Add(_browserFallback);
            transcript.Controls.Add(_empty);

            Panel tabBar = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = SystemColors.Control };
            // 纵向 11 把 9 磅标题撑满 40 像素条，页签页从 42 起，不会露在标题下面。
            _tabs.ImageList = _conversationTabIcons;
            _tabs.Padding = new Point(6, 11);
            _tabs.Dock = DockStyle.Fill;
            tabBar.Controls.Add(_tabs);
            tabBar.Controls.Add(TabIconButton("icon_session_add.png", AppPresentation.AiNewConversationTip, (sender, args) => StartConversation("新对话", string.Empty, false)));
            Button historyButton = TabIconButton("icon_session_history.png", AppPresentation.AiArchivedTip, (sender, args) => ShowHistoryMenu((Control)sender));
            tabBar.Controls.Add(historyButton);
            if (_closePanel != null)
                tabBar.Controls.Add(TabIconButton("icon_close.png", AppPresentation.AiCloseChatTip, (sender, args) => _closePanel()));
            Controls.Add(transcript);
            Controls.Add(line);
            Controls.Add(_composer);
            Controls.Add(tabBar);

            _assistant.Conversations.Changed += ConversationsChanged;
            _assistant.Models.Changed += ChoicesChanged;
            _assistant.Skills.Changed += ChoicesChanged;
            ReloadChoices();
            ReloadTabs();
            RefreshSqlReference();
            LayoutComposer();
        }

        /// <summary>没有对话时新建一条，已有对话时保持当前选择。</summary>
        public void EnsureConversation()
        {
            if (_assistant.Conversations.GetAll().Count == 0)
            {
                StartConversation("新对话", string.Empty, false);
                return;
            }

            if (string.IsNullOrEmpty(_currentId))
                SelectFirstActive();
        }

        /// <summary>另建一条对话并填入草稿。报错入口用它，不切走查询窗里正在写的那条。</summary>
        /// <param name="title">标题。报错入口传入「SQL 报错」，并锁定标题。</param>
        /// <param name="draft">放进输入框的文字。不自动发送。</param>
        public void StartConversation(string title, string draft)
        {
            StartConversation(title, draft, true);
        }

        /// <summary>查询页签切换或标题变化时刷新输入框上方的引用。</summary>
        /// <remarks>用户删掉引用后，这次编写期间保持不附带，直到这次发送结束。</remarks>
        public void NotifyEditorChanged()
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(NotifyEditorChanged));
                return;
            }

            RefreshSqlReference();
        }

        /// <summary>刷新模型和 Skill 菜单，尽量保持原来的选择。</summary>
        public void ReloadChoices()
        {
            if (IsDisposed) return;
            _models.Clear();
            string stored = _settings.Get(CoreSettings.AiModelChoice) ?? string.Empty;
            int bar = stored.IndexOf('|');
            string storedProfile = bar > 0 ? stored.Substring(0, bar) : string.Empty;
            string storedModel = bar >= 0 ? stored.Substring(bar + 1) : string.Empty;
            ModelPick preferred = null;
            ModelPick defaultPick = null;
            foreach (AiModelProfile profile in _assistant.Models.GetAll())
            {
                foreach (string modelId in AiModelCatalog.Ids(profile))
                {
                    ModelPick pick = new ModelPick(profile, modelId);
                    _models.Add(pick);
                    if (preferred == null
                        && string.Equals(profile.Id, storedProfile, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(modelId, storedModel, StringComparison.OrdinalIgnoreCase))
                        preferred = pick;
                    if (defaultPick == null && profile.IsDefault)
                        defaultPick = pick;
                }
            }

            ModelPick chosen = preferred ?? defaultPick ?? (_models.Count == 0 ? null : _models[0]);
            _modelId = chosen == null ? null : chosen.Profile.Id;
            _modelName = chosen == null ? null : chosen.ModelId;
            string nextChoice = string.IsNullOrEmpty(_modelId) ? string.Empty : _modelId + "|" + (_modelName ?? string.Empty);
            if (nextChoice.Length > 0 && !string.Equals(stored, nextChoice, StringComparison.Ordinal))
                _settings.Set(CoreSettings.AiModelChoice, nextChoice);
            _skills.Clear();
            _skills.Add(new SkillPick(null));
            foreach (AiSkill skill in _assistant.Skills.GetAll())
                _skills.Add(new SkillPick(skill));
            LoadSkillChoice();
            PruneSkillIds();
            FillChoiceMenu(_modelMenu, true);
            FillChoiceMenu(_skillMenu, false);
            ApplyChoiceButtons();
        }

        /// <summary>释放时退订存储事件并关掉浏览器，避免查询窗关掉后仍被回调。</summary>
        /// <param name="disposing">为 true 时释放托管资源。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _assistant.Conversations.Changed -= ConversationsChanged;
                _assistant.Models.Changed -= ChoicesChanged;
                _assistant.Skills.Changed -= ChoicesChanged;
                _tips.Dispose();
                _view.Dispose();
                _tabMenu.Dispose();
                _modelMenu.Dispose();
                _skillMenu.Dispose();
                _historyMenu.Dispose();
                foreach (Image icon in _tabIcons)
                    icon.Dispose();
                _tabIcons.Clear();
                foreach (Image icon in _mentionIcons.Values)
                    icon.Dispose();
                _mentionIcons.Clear();
                _tabs.ImageList = null;
                _conversationTabIcons.Dispose();
                _addIcon.Dispose();
                _sendIcon.Dispose();
                _sendIconGray.Dispose();
                _chipFont.Dispose();
                if (_mention != null)
                    _mention.Dispose();
            }

            base.Dispose(disposing);
        }

        /// <summary>句柄创建后再初始化浏览器，此时控件已经挂到窗体上。</summary>
        /// <param name="e">句柄参数。</param>
        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            if (IsHandleCreated)
                BeginInvoke(new Action(BeginBrowser));
        }

        private async void BeginBrowser()
        {
            if (_browserStarted || IsDisposed) return;
            _browserStarted = true;
            try
            {
                CoreWebView2Environment environment = await AiChatBrowserHost.CreateAsync(_assistant.BrowserDataDirectory);
                if (IsDisposed) return;
                await _view.EnsureCoreWebView2Async(environment);
                if (IsDisposed || _view.CoreWebView2 == null) return;
                CoreWebView2Settings settings = _view.CoreWebView2.Settings;
                settings.AreDefaultScriptDialogsEnabled = false;
                settings.IsWebMessageEnabled = false;
                settings.AreHostObjectsAllowed = false;
                settings.AreDevToolsEnabled = false;
                settings.IsStatusBarEnabled = false;
                _view.CoreWebView2.NewWindowRequested += (sender, args) => { args.Handled = true; };
                _view.CoreWebView2.NavigationStarting += (sender, args) =>
                {
                    if (TryOpenSent(args.Uri))
                    {
                        args.Cancel = true;
                        return;
                    }

                    if (args.IsUserInitiated)
                        args.Cancel = true;
                };
                _browserReady = true;
                ApplyTranscriptVisibility();
                ShowTranscript();
            }
            catch (Exception exception)
            {
                if (IsDisposed) return;
                _browserFailed = true;
                _browserFallback.Text = "缺少 WebView2 运行库，对话无法显示排版。";
                ApplyTranscriptVisibility();
                ExceptionDetailForm.Show(FindForm(), exception);
            }
        }

        private void StartConversation(string title, string draft, bool titleLocked)
        {
            try
            {
                AiConversation created = _assistant.Conversations.Create(title, titleLocked);
                _assistant.WriteAction("conversation.create", "title=" + created.Title, created.Id, created.Title);
                _omitSql = false;
                ReloadTabs();
                Select(created.Id);
                _input.Text = draft ?? string.Empty;
                RefreshSqlReference();
                _input.Focus();
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(FindForm(), exception);
            }
        }

        private void Guard(Action action)
        {
            try
            {
                action();
            }
            catch (Exception exception)
            {
                ExceptionDetailForm.Show(FindForm(), exception);
            }
        }

        private void RenameCurrent()
        {
            AiConversation conversation = Current();
            if (conversation == null) return;
            if (!TextPromptDialog.TryShow(FindForm(), "重命名对话", "请输入对话标题", conversation.Title, out string title))
                return;
            _assistant.Conversations.Rename(conversation.Id, title);
            _assistant.WriteAction("conversation.rename", "title=" + title, conversation.Id, title);
        }

        private void ArchiveCurrent()
        {
            AiConversation conversation = Current();
            if (conversation == null) return;
            ClearConversationAlerts(conversation.Id);
            if (string.Equals(_currentId, conversation.Id, StringComparison.OrdinalIgnoreCase))
                _currentId = null;
            _assistant.Conversations.SetArchived(conversation.Id, true);
            _assistant.WriteAction("conversation.archive", "title=" + conversation.Title, conversation.Id, conversation.Title);
        }

        private void DeleteCurrent()
        {
            AiConversation conversation = Current();
            if (conversation == null) return;
            DeleteConversation(conversation, true);
        }

        private void DeleteConversation(AiConversation conversation, bool confirmWhenContent)
        {
            int count = conversation.Messages == null ? 0 : conversation.Messages.Count;
            if (confirmWhenContent && count > 0
                && MessageBox.Show(FindForm(), "删除这条对话？已经发出的内容仍留在审计日志里。", "删除对话", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                return;
            if (string.Equals(_currentId, conversation.Id, StringComparison.OrdinalIgnoreCase))
                _currentId = null;
            ClearConversationAlerts(conversation.Id);
            _assistant.Conversations.Delete(conversation.Id);
            _assistant.WriteAction(
                "conversation.delete",
                "title=" + conversation.Title + "\r\nmessages=" + count.ToString(),
                conversation.Id,
                conversation.Title);
            SelectFirstActive();
        }

        private void RestoreConversation(string id)
        {
            AiConversation conversation = _assistant.Conversations.Get(id);
            if (conversation == null) return;
            _currentId = conversation.Id;
            _assistant.Conversations.SetArchived(conversation.Id, false);
            _assistant.WriteAction("conversation.restore", "title=" + conversation.Title, conversation.Id, conversation.Title);
            Select(conversation.Id);
        }

        private async void SendButtonClick(object sender, EventArgs e)
        {
            if (_sending) return;
            string raw = _input.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw))
            {
                MessageBox.Show(FindForm(), "请输入问题。", "AI 对话", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ModelPick model = CurrentModel();
            if (model == null)
            {
                MessageBox.Show(FindForm(), "请先在「设置 → AI 设置 → 模型设置」中保存一个模型。", "AI 对话", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            AiConversation conversation = Current();
            if (conversation == null)
            {
                MessageBox.Show(FindForm(), "请先新建对话。", "AI 对话", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            AiEditorSnapshot snapshot = _readEditor() ?? new AiEditorSnapshot();
            bool attachSql = !_omitSql && snapshot.HasTab;
            string conversationId = conversation.Id;
            IList<AiChatTurn> history = conversation.Messages;
            List<Attachment> sentAttachments = new List<Attachment>(_attachments);
            _sending = true;
            _errorConversationIds.Remove(conversationId);
            _unreadConversationIds.Remove(conversationId);
            SetComposerBusy(true);
            HideMention();
            _input.Clear();
            _attachments.Clear();
            RebuildAttachmentChips();
            _pendingConversationId = conversationId;
            _pendingUser = raw;
            RefreshConversationTabIcons();
            ShowTranscript();
            try
            {
                string userNote;
                string sqlNote = "无";
                string safeUser = AiTextRedactor.Redact(raw, out userNote);
                string safeSql = string.Empty;
                if (attachSql)
                    safeSql = AiTextRedactor.Redact(snapshot.Sql, out sqlNote);
                string skillNote = "无";
                StringBuilder skillNames = new StringBuilder();
                StringBuilder skillBodies = new StringBuilder();
                List<AiSentSection> sent = new List<AiSentSection>();
                sent.Add(new AiSentSection
                {
                    Kind = AiSentSection.Question,
                    Title = AppPresentation.AiSentQuestion,
                    Body = raw
                });
                string systemBase = AiPrompt.SystemText(model.ModelId);
                sent.Add(new AiSentSection
                {
                    Kind = AiSentSection.System,
                    Title = AppPresentation.AiSentSystem,
                    Body = systemBase
                });
                StringBuilder systemText = new StringBuilder(systemBase);
                foreach (SkillPick pick in SelectedSkills())
                {
                    if (pick.Skill == null || string.IsNullOrWhiteSpace(pick.Skill.Body)) continue;
                    sent.Add(new AiSentSection
                    {
                        Kind = AiSentSection.Skill,
                        Title = pick.Skill.Name ?? string.Empty,
                        Body = pick.Skill.Body
                    });
                    string safeSkill = AiTextRedactor.Redact(pick.Skill.Body, out string note);
                    skillNote = CombineNotes(skillNote, note);
                    if (skillNames.Length > 0) skillNames.Append("、");
                    skillNames.Append(pick.Skill.Name ?? string.Empty);
                    if (skillBodies.Length > 0)
                        skillBodies.AppendLine().AppendLine();
                    skillBodies.Append(safeSkill);
                    systemText.Append("\r\n\r\n");
                    systemText.Append(AiPrompt.Skill(pick.Skill.Name, safeSkill));
                }

                string system = systemText.ToString();
                string dialect = AiTextRedactor.DescribeDialect(snapshot.ProviderId);
                sent.Add(new AiSentSection
                {
                    Kind = AiSentSection.Dialect,
                    Title = AppPresentation.AiSentDialect,
                    Body = dialect
                });
                if (attachSql)
                {
                    sent.Add(new AiSentSection
                    {
                        Kind = AiSentSection.Tab,
                        Title = snapshot.TabName ?? string.Empty,
                        Body = snapshot.Sql ?? string.Empty
                    });
                }

                string contextNote = "无";
                bool includesPreview = false;
                StringBuilder context = new StringBuilder();
                foreach (Attachment attachment in sentAttachments)
                {
                    sent.Add(new AiSentSection
                    {
                        Kind = AiSentSection.Citation,
                        Title = attachment.Label ?? string.Empty,
                        Body = attachment.Body ?? string.Empty
                    });
                    string safeBody = AiTextRedactor.Redact(attachment.Body, out string note);
                    contextNote = CombineNotes(contextNote, note);
                    if (attachment.Preview)
                        includesPreview = true;
                    if (context.Length > 0)
                        context.AppendLine().AppendLine();
                    context.Append(AiPrompt.Context(attachment.Label, safeBody));
                }

                string historyText = AiAssistant.FormatHistory(history);
                if (historyText.Length > 0)
                {
                    sent.Add(new AiSentSection
                    {
                        Kind = AiSentSection.History,
                        Title = AppPresentation.AiSentHistory,
                        Body = historyText
                    });
                }

                AiSentCall call = new AiSentCall
                {
                    Url = AiChatClient.ChatUrl(
                        AiProtocolNames.Parse(model.Profile.Protocol),
                        model.Profile.Endpoint,
                        model.ModelId),
                    Protocol = model.Profile.Protocol,
                    ModelDisplayName = model.Profile.DisplayName,
                    ModelId = model.ModelId,
                    SentUtc = DateTime.UtcNow.ToString("o")
                };
                _pendingSent = sent;
                _pendingCall = call;

                string finalUser = AiAssistant.ComposeUserMessage(
                    dialect,
                    attachSql,
                    attachSql ? snapshot.TabName : string.Empty,
                    safeSql,
                    safeUser,
                    context.ToString());
                string apiKey = AiSecretProtector.Unprotect(model.Profile.ProtectedKey);
                AiAuditEntry entry = new AiAuditEntry
                {
                    Action = "chat.send",
                    ConversationId = conversationId,
                    ConversationTitle = conversation.Title,
                    ModelDisplayName = AiModelCaption.Audit(model.Profile.DisplayName, model.ModelId),
                    Protocol = model.Profile.Protocol,
                    Endpoint = AiEndpoints.Describe(model.Profile.Endpoint),
                    SkillName = skillNames.ToString(),
                    SkillBody = skillBodies.ToString(),
                    SystemText = system,
                    UserText = safeUser,
                    HistoryText = AiAssistant.FormatHistory(history),
                    TabName = attachSql ? snapshot.TabName : string.Empty,
                    TabSql = attachSql ? safeSql : string.Empty,
                    HasTab = attachSql,
                    Dialect = dialect,
                    ContextText = context.ToString(),
                    IncludesPreview = includesPreview,
                    Redaction = CombineNotes(userNote, CombineNotes(sqlNote, CombineNotes(skillNote, contextNote)))
                };
                _pendingUser = safeUser;
                ShowTranscript();
                AiReply reply = await _assistant.AskAsync(
                    model.Profile,
                    apiKey,
                    model.ModelId,
                    entry,
                    history,
                    system,
                    finalUser);
                RunOnUiThread(() =>
                {
                    if (IsDisposed) return;
                    call.HttpStatus = reply.HttpStatus;
                    call.ElapsedMs = reply.ElapsedMs;
                    _pendingUser = null;
                    _pendingConversationId = null;
                    _pendingSent = null;
                    _pendingCall = null;
                    _assistant.Conversations.Append(conversationId, new AiChatTurn
                    {
                        Role = "user",
                        Text = safeUser,
                        Sent = sent,
                        Call = call
                    });
                    _assistant.Conversations.Append(conversationId, new AiChatTurn
                    {
                        Role = "assistant",
                        Text = reply.Text,
                        InputTokens = reply.InputTokens,
                        OutputTokens = reply.OutputTokens
                    });
                    if (string.Equals(_currentId, conversationId, StringComparison.OrdinalIgnoreCase))
                        ClearConversationAlerts(conversationId);
                    else
                        MarkConversationUnread(conversationId);
                    RefreshConversationTabIcons();
                    ShowTranscript();
                    MaybeTitle(conversationId, model, apiKey);
                });
            }
            catch (Exception exception)
            {
                RunOnUiThread(() =>
                {
                    if (IsDisposed) return;
                    _pendingUser = null;
                    _pendingConversationId = null;
                    _pendingSent = null;
                    _pendingCall = null;
                    MarkConversationError(conversationId);
                    if (string.Equals(_currentId, conversationId, StringComparison.OrdinalIgnoreCase)
                        && string.IsNullOrEmpty(_input.Text))
                        _input.Text = raw;
                    if (_attachments.Count == 0)
                    {
                        foreach (Attachment item in sentAttachments)
                            AddAttachment(item.Label, item.Body, item.Kind);
                    }
                    RefreshConversationTabIcons();
                    ShowTranscript();
                    ExceptionDetailForm.Show(FindForm(), exception);
                });
            }
            finally
            {
                _sending = false;
                RunOnUiThread(() =>
                {
                    if (IsDisposed) return;
                    SetComposerBusy(false);
                    _omitSql = false;
                    RefreshSqlReference();
                    RefreshConversationTabIcons();
                });
            }
        }

        private void ConversationsChanged(object sender, EventArgs e)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ReloadFromStore));
                return;
            }

            ReloadFromStore();
        }

        private void ChoicesChanged(object sender, EventArgs e)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                BeginInvoke(new Action(ReloadChoices));
                return;
            }

            ReloadChoices();
        }

        private void ReloadFromStore()
        {
            string id = _currentId;
            ReloadTabs();
            if (!string.IsNullOrEmpty(id))
                Select(id);
            if (string.IsNullOrEmpty(_currentId) || !TabExists(_currentId))
                SelectFirstActive();
            ShowTranscript();
        }

        private void ReloadTabs()
        {
            _binding = true;
            _tabs.TabPages.Clear();
            foreach (AiConversation conversation in _assistant.Conversations.GetAll())
            {
                if (conversation == null || conversation.Archived) continue;
                _tabs.TabPages.Add(new TabPage(CloseMark.Caption(conversation.Title, _tabs.Font))
                {
                    Tag = conversation.Id,
                    ImageIndex = ConversationTabIconDefault
                });
            }

            RefreshConversationTabIcons();
            ApplyTranscriptVisibility();
            _binding = false;
        }

        /// <summary>按会话状态刷新所有页签图标。状态优先级是等待、失败、未读、默认。</summary>
        private void RefreshConversationTabIcons()
        {
            for (int index = 0; index < _tabs.TabPages.Count; index++)
            {
                TabPage page = _tabs.TabPages[index];
                page.ImageIndex = ResolveConversationTabIcon(page.Tag as string);
            }
        }

        /// <summary>计算某条会话应显示的图标索引。</summary>
        /// <param name="conversationId">会话标识。</param>
        /// <returns>页签图标索引。</returns>
        private int ResolveConversationTabIcon(string conversationId)
        {
            if (!string.IsNullOrEmpty(conversationId)
                && _sending
                && string.Equals(conversationId, _pendingConversationId, StringComparison.OrdinalIgnoreCase))
                return ConversationTabIconPending;
            if (!string.IsNullOrEmpty(conversationId) && _errorConversationIds.Contains(conversationId))
                return ConversationTabIconError;
            if (!string.IsNullOrEmpty(conversationId) && _unreadConversationIds.Contains(conversationId))
                return ConversationTabIconUnread;
            return ConversationTabIconDefault;
        }

        /// <summary>没有对话时用整块空白盖住正文，避免 WebView 吃掉点击。</summary>
        private void ApplyTranscriptVisibility()
        {
            bool any = _tabs.TabPages.Count > 0;
            _tabs.Visible = any;
            _empty.Visible = !any;
            if (!any)
            {
                _view.Visible = false;
                _browserFallback.Visible = false;
                CenterEmptyButton();
                return;
            }

            _browserFallback.Visible = _browserFailed;
            _view.Visible = !_browserFailed;
        }

        private void TabsSelectedIndexChanged(object sender, EventArgs e)
        {
            if (_binding) return;
            if (_tabs.SelectedTab == null) return;
            _currentId = _tabs.SelectedTab.Tag as string;
            ClearConversationAlerts(_currentId);
            RefreshConversationTabIcons();
            ShowTranscript();
        }

        private void TabsMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            int index = TabIndexAt(e.Location);
            if (index < 0) return;
            if (!CloseMark.Hit(_tabs.GetTabRect(index), e.Location)) return;
            string id = _tabs.TabPages[index].Tag as string;
            BeginInvoke(new Action(() => Guard(() => CloseByGlyph(id))));
        }

        private void TabsMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            int index = TabIndexAt(e.Location);
            if (index >= 0 && _tabs.SelectedIndex != index)
                _tabs.SelectedIndex = index;
            bool onTab = index >= 0;
            _renameItem.Enabled = onTab;
            _archiveItem.Enabled = onTab;
            _deleteItem.Enabled = onTab;
            FillArchivedMenu();
            _tabMenu.Show(_tabs, e.Location);
        }

        private void EmptyMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            _renameItem.Enabled = false;
            _archiveItem.Enabled = false;
            _deleteItem.Enabled = false;
            FillArchivedMenu();
            Control source = sender as Control ?? _empty;
            _tabMenu.Show(source, e.Location);
        }

        /// <summary>把「请新建对话」放在正文区正中。按钮本身没有边框。</summary>
        private void CenterEmptyButton()
        {
            if (_emptyButton == null || _empty == null) return;
            Size size = _emptyButton.PreferredSize;
            if (_emptyButton.Width > 0) size.Width = _emptyButton.Width;
            if (_emptyButton.Height > 0) size.Height = _emptyButton.Height;
            _emptyButton.Location = new Point(
                Math.Max(0, (_empty.ClientSize.Width - size.Width) / 2),
                Math.Max(0, (_empty.ClientSize.Height - size.Height) / 2));
        }

        /// <summary>单独按 Enter 发送。Shift+Enter 留给文本框换行。等待回复时不重复发送。</summary>
        private void InputKeyDown(object sender, KeyEventArgs e)
        {
            if (_mention != null && _mention.Visible)
            {
                if (e.KeyCode == Keys.Escape)
                {
                    HideMention();
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.Down)
                {
                    _mention.MoveSelection(1);
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.Up)
                {
                    _mention.MoveSelection(-1);
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    return;
                }

                if (e.KeyCode == Keys.Enter && !e.Shift)
                {
                    _mention.AcceptCurrent();
                    e.SuppressKeyPress = true;
                    e.Handled = true;
                    return;
                }
            }

            if (e.KeyCode != Keys.Enter || e.Shift || e.Control || e.Alt) return;
            e.SuppressKeyPress = true;
            e.Handled = true;
            if (_sending) return;
            SendButtonClick(_send, EventArgs.Empty);
        }

        private void InputTextChanged(object sender, EventArgs e)
        {
            if (_suppressMention || _sending || _mention == null) return;
            string text = _input.Text ?? string.Empty;
            int caret = _input.SelectionStart;
            string filter;
            int at = MentionAt(text, caret, out filter);
            if (at < 0)
            {
                HideMention();
                return;
            }

            _mentionStart = at;
            ShowMention(filter);
        }

        private static int MentionAt(string text, int caret, out string filter)
        {
            filter = string.Empty;
            if (text == null || caret < 1 || caret > text.Length) return -1;
            int start = caret;
            while (start > 0 && !char.IsWhiteSpace(text[start - 1]))
                start--;
            if (start >= text.Length || text[start] != '@') return -1;
            filter = text.Substring(start + 1, caret - start - 1);
            return start;
        }

        private void ShowMention(string filter)
        {
            AiMentionSource source = _readMentions() ?? new AiMentionSource();
            List<MentionChoice> choices = new List<MentionChoice>();
            choices.Add(new MentionChoice("文件", MentionKind.File, null, null));
            if (source.HasPreview && !string.IsNullOrWhiteSpace(source.PreviewText))
                choices.Add(new MentionChoice("结果预览", MentionKind.Preview, source.PreviewLabel, source.PreviewText));
            else
                choices.Add(new MentionChoice("结果预览", MentionKind.Preview, null, null) { Enabled = false });
            IReadOnlyList<AiSavedQuery> queries = source.Queries ?? new AiSavedQuery[0];
            foreach (AiSavedQuery query in queries)
            {
                if (query == null || string.IsNullOrWhiteSpace(query.Name)) continue;
                if (!string.IsNullOrEmpty(filter)
                    && query.Name.IndexOf(filter, StringComparison.CurrentCultureIgnoreCase) < 0)
                    continue;
                choices.Add(new MentionChoice(query.Name, MentionKind.Query, query.Name, query.Sql ?? string.Empty));
            }

            Point origin = _input.GetPositionFromCharIndex(Math.Max(0, _input.SelectionStart));
            Point screen = _input.PointToScreen(new Point(origin.X, origin.Y + _input.Font.Height));
            _mention.ShowChoices(choices, screen, FindForm());
        }

        private void HideMention()
        {
            if (_mention != null && _mention.Visible)
                _mention.Hide();
        }

        private void AcceptMention(MentionChoice choice)
        {
            HideMention();
            if (choice == null || !choice.Enabled) return;
            if (choice.Kind == MentionKind.File)
            {
                AttachFile();
                return;
            }

            if (choice.Kind == MentionKind.Preview && string.IsNullOrWhiteSpace(choice.Body))
                return;
            RemoveMentionToken();
            string label = choice.Kind == MentionKind.Preview
                ? "预览 " + (string.IsNullOrWhiteSpace(choice.Label) ? "结果" : choice.Label)
                : choice.Label;
            AddAttachment(label, choice.Body ?? string.Empty, choice.Kind);
        }

        private void AttachFile()
        {
            RemoveMentionToken();
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Title = "选择要引用的文件";
                dialog.Filter = "文本文件|*.txt;*.sql;*.csv;*.json;*.md;*.log|所有文件|*.*";
                if (dialog.ShowDialog(FindForm()) != DialogResult.OK) return;
                try
                {
                    string body = ReadTextFile(dialog.FileName, out bool truncated);
                    if (truncated)
                        body += Environment.NewLine + "（文件已截断，只附上前 64KB。）";
                    AddAttachment(Path.GetFileName(dialog.FileName), body, MentionKind.File);
                }
                catch (Exception exception)
                {
                    ExceptionDetailForm.Show(FindForm(), exception);
                }
            }
        }

        private static string ReadTextFile(string path, out bool truncated)
        {
            truncated = false;
            byte[] buffer = new byte[MaxFileBytes + 1];
            int read;
            using (FileStream stream = File.OpenRead(path))
                read = stream.Read(buffer, 0, buffer.Length);
            int length = Math.Min(read, MaxFileBytes);
            for (int index = 0; index < length; index++)
            {
                if (buffer[index] == 0)
                    throw new InvalidOperationException("这个文件不是文本，不能作为引用。");
            }

            truncated = read > MaxFileBytes;
            return Encoding.UTF8.GetString(buffer, 0, length);
        }

        private void RemoveMentionToken()
        {
            string text = _input.Text ?? string.Empty;
            int caret = _input.SelectionStart;
            if (_mentionStart < 0 || _mentionStart > caret || _mentionStart >= text.Length) return;
            _suppressMention = true;
            _input.Text = text.Remove(_mentionStart, caret - _mentionStart);
            _input.SelectionStart = Math.Min(_mentionStart, _input.Text.Length);
            _suppressMention = false;
            _mentionStart = -1;
        }

        private void AddAttachment(string label, string body, MentionKind kind)
        {
            Attachment attachment = new Attachment
            {
                Label = label ?? string.Empty,
                Body = body ?? string.Empty,
                Kind = kind
            };
            Button chip = ReferenceChip(kind);
            chip.Text = CloseMark.Caption(attachment.Label, chip.Font);
            attachment.View = chip;
            chip.MouseDown += (sender, args) =>
            {
                if (args.Button != MouseButtons.Left) return;
                if (!CloseMark.Hit(chip.ClientRectangle, args.Location)) return;
                _attachments.Remove(attachment);
                BeginInvoke(new Action(RebuildAttachmentChips));
            };
            _tips.SetToolTip(chip, kind == MentionKind.Preview ? AppPresentation.AiPreviewChipTip : AppPresentation.AiAttachmentChipTip);
            _attachments.Add(attachment);
            _chipRow.Controls.Add(chip);
            LayoutComposer();
        }

        private void RebuildAttachmentChips()
        {
            List<Control> keep = new List<Control>();
            foreach (Attachment attachment in _attachments)
            {
                if (attachment.View != null)
                    keep.Add(attachment.View);
            }

            for (int index = _chipRow.Controls.Count - 1; index >= 0; index--)
            {
                Control control = _chipRow.Controls[index];
                if (control == _skillButton || control == _sqlRef) continue;
                _chipRow.Controls.RemoveAt(index);
                if (!keep.Contains(control))
                    control.Dispose();
            }

            foreach (Control view in keep)
            {
                if (!view.IsDisposed)
                    _chipRow.Controls.Add(view);
            }

            LayoutComposer();
        }

        /// <summary>按当前宽度重算引用行、灰色说明和模型行的高度，再把底栏撑到刚好放下。</summary>
        /// <remarks>说明在创建时已抽定，这里只换行，不改文字。测量比文字区略窄并多留几像素，避免末行被挡住。宽度还没有时不改。</remarks>
        private void LayoutComposer()
        {
            if (_composer == null || _chipRow == null || _mentionHint == null || _modelRow == null) return;
            int width = _composer.ClientSize.Width;
            if (width <= 0) return;
            Size preferred = _chipRow.GetPreferredSize(new Size(width, 0));
            int chips = Math.Max(36, preferred.Height);
            if (_chipRow.Height != chips)
                _chipRow.Height = chips;
            int hintWidth = Math.Max(1, width - _mentionHint.Padding.Horizontal - 4);
            Size measured = TextRenderer.MeasureText(
                _mentionHint.Text,
                _mentionHint.Font,
                new Size(hintWidth, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.Left | TextFormatFlags.TextBoxControl);
            int hint = Math.Max(20, measured.Height + _mentionHint.Padding.Vertical + 4);
            if (_mentionHint.Height != hint)
                _mentionHint.Height = hint;
            Size sendSize = _send.GetPreferredSize(Size.Empty);
            int sendWidth = SendButtonWidth();
            if (_send.Width != sendWidth)
                _send.Width = sendWidth;
            int modelRowHeight = Math.Max(36, sendSize.Height);
            if (_modelRow.Height != modelRowHeight)
                _modelRow.Height = modelRowHeight;
            int modelY = Math.Max(0, (_modelRow.ClientSize.Height - _modelButton.Height) / 2);
            if (_modelButton.Left != 8 || _modelButton.Top != modelY)
                _modelButton.Location = new Point(8, modelY);
            int next = _modelRow.Height + hint + 72 + chips + 8;
            if (_composer.Height != next)
                _composer.Height = next;
        }

        /// <summary>发送按钮的宽度。系统首选宽度不含左侧图标，按图标、文字和左右空隙另算。</summary>
        /// <returns>能放下 24 像素图标和「发送」两字、两侧仍有空隙的像素宽度。</returns>
        private int SendButtonWidth()
        {
            int textWidth = TextRenderer.MeasureText(_send.Text, _send.Font).Width;
            int icon = _sendIcon == null ? 0 : _sendIcon.Width;
            return textWidth + icon + _send.Padding.Horizontal + 8;
        }

        /// <summary>等待回复时禁用输入和发送，并把发送图标换成灰色。</summary>
        /// <param name="busy">为 true 时正在等模型回复。</param>
        /// <remarks>只在界面线程调用。回复到达或失败后传 false，输入框恢复可编辑。</remarks>
        private void SetComposerBusy(bool busy)
        {
            _input.Enabled = !busy;
            _send.Enabled = !busy;
            _send.Image = busy ? _sendIconGray : _sendIcon;
        }

        /// <summary>只有点中引用文字末尾的关闭标记才去掉这次页签。</summary>
        private void SqlRefMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            if (!CloseMark.Hit(_sqlRef.ClientRectangle, e.Location)) return;
            _omitSql = true;
            RefreshSqlReference();
        }

        private void CloseByGlyph(string id)
        {
            AiConversation conversation = _assistant.Conversations.Get(id);
            if (conversation == null) return;
            int count = conversation.Messages == null ? 0 : conversation.Messages.Count;
            if (count == 0 && string.Equals(conversation.Title, "新对话", StringComparison.Ordinal))
            {
                DeleteConversation(conversation, false);
                return;
            }

            if (!string.Equals(_currentId, conversation.Id, StringComparison.OrdinalIgnoreCase))
                Select(conversation.Id);
            ArchiveCurrent();
        }

        private void FillArchivedMenu()
        {
            FillArchivedInto(_archivedMenu.DropDownItems);
            _archivedMenu.Enabled = _archivedMenu.DropDownItems.Count > 0;
        }

        /// <summary>点历史按钮时弹出已归档对话。没有归档时留一条不可用说明。</summary>
        /// <param name="anchor">菜单对齐的按钮。</param>
        private void ShowHistoryMenu(Control anchor)
        {
            if (anchor == null) return;
            _historyMenu.Items.Clear();
            FillArchivedInto(_historyMenu.Items);
            if (_historyMenu.Items.Count == 0)
                _historyMenu.Items.Add(new ToolStripMenuItem(AppPresentation.AiNoArchivedConversation) { Enabled = false });
            _historyMenu.Show(anchor, new Point(0, anchor.Height));
        }

        /// <summary>把已归档对话填进菜单。右键「已归档」和历史按钮共用这一份。</summary>
        /// <param name="items">要清空后重填的菜单项。</param>
        private void FillArchivedInto(ToolStripItemCollection items)
        {
            if (items == null) return;
            items.Clear();
            foreach (AiConversation conversation in _assistant.Conversations.GetAll())
            {
                if (conversation == null || !conversation.Archived) continue;
                string id = conversation.Id;
                items.Add(conversation.Title, null, (sender, args) => Guard(() => RestoreConversation(id)));
            }
        }

        /// <summary>页签条右侧的图标按钮。图像由本控件在释放时丢掉。</summary>
        /// <param name="fileName">Resources 下的文件名。</param>
        /// <param name="tip">悬浮说明。</param>
        /// <param name="click">点击动作。</param>
        /// <returns>停靠在右侧、宽 40 的平按钮。高度随 40 像素页签条，点按区域是正方形。</returns>
        private Button TabIconButton(string fileName, string tip, EventHandler click)
        {
            Image image = DatabaseTreeImageCatalog.LoadActionIcon(fileName, 24);
            _tabIcons.Add(image);
            Button button = new Button
            {
                Dock = DockStyle.Right,
                Width = 40,
                Image = image,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Control,
                Cursor = Cursors.Hand,
                Text = string.Empty,
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = SystemColors.ControlDark;
            button.FlatAppearance.MouseDownBackColor = SystemColors.ControlDarkDark;
            button.Click += click;
            _tips.SetToolTip(button, tip);
            return button;
        }

        /// <summary>认出查看请求详情的链接并打开只读窗。其他地址返回 false。</summary>
        /// <param name="uriText">WebView 将要打开的地址。</param>
        /// <returns>这是请求详情链接时为 true。调用方应取消跳转。</returns>
        private bool TryOpenSent(string uriText)
        {
            if (!Uri.TryCreate(uriText ?? string.Empty, UriKind.Absolute, out Uri uri))
                return false;
            if (!string.Equals(uri.Host, "db2sheet.invalid", StringComparison.OrdinalIgnoreCase))
                return false;
            string path = uri.AbsolutePath ?? string.Empty;
            const string prefix = "/sent/";
            if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return false;
            string key = path.Substring(prefix.Length).Trim('/');
            if (!IsDisposed && IsHandleCreated)
                BeginInvoke(new Action(() => ShowSent(key)));
            return true;
        }

        /// <summary>按序号或等待中的这一轮打开请求详情。已有回复时带上那一次的输入 token 和调用信息。</summary>
        /// <param name="key">pending，或往来列表里的序号。</param>
        private void ShowSent(string key)
        {
            if (IsDisposed) return;
            IList<AiSentSection> sent = null;
            AiSentCall call = null;
            int? inputTokens = null;
            if (string.Equals(key, "pending", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(_currentId, _pendingConversationId, StringComparison.OrdinalIgnoreCase))
                {
                    sent = _pendingSent;
                    call = _pendingCall;
                }
            }
            else
            {
                int index;
                AiConversation conversation = Current();
                if (int.TryParse(key, NumberStyles.None, CultureInfo.InvariantCulture, out index)
                    && conversation != null
                    && conversation.Messages != null
                    && index >= 0
                    && index < conversation.Messages.Count
                    && conversation.Messages[index] != null)
                {
                    sent = conversation.Messages[index].Sent;
                    call = conversation.Messages[index].Call;
                    if (index + 1 < conversation.Messages.Count)
                    {
                        AiChatTurn reply = conversation.Messages[index + 1];
                        if (reply != null
                            && string.Equals(reply.Role, "assistant", StringComparison.OrdinalIgnoreCase)
                            && reply.InputTokens.HasValue)
                        {
                            inputTokens = reply.InputTokens;
                        }
                    }
                }
            }

            if (sent == null || sent.Count == 0) return;
            AiSentDetailForm.Show(FindForm(), sent, inputTokens, call, _assistant.BrowserDataDirectory);
        }

        /// <summary>第 1、3、7、13 次用户消息后，另起一次短请求生成标题。之后不再改。</summary>
        /// <param name="conversationId">刚写入回复的对话。</param>
        /// <param name="model">这次回答用的型号。</param>
        /// <param name="apiKey">已经解开的密钥。</param>
        private void MaybeTitle(string conversationId, ModelPick model, string apiKey)
        {
            AiConversation conversation = _assistant.Conversations.Get(conversationId);
            if (!ShouldRefreshTitle(conversation) || model == null) return;
            int serial = ++_titleSerial;
            BeginTitle(serial, conversation, model, apiKey);
        }

        /// <summary>标题请求失败时保留原标题，不弹错误详情。用户的问题已经成功。</summary>
        /// <param name="serial">这次请求的序号。更新的请求开始后，旧结果不再写标题。</param>
        /// <param name="conversation">发起时的对话副本，含刚写入的问答。</param>
        /// <param name="model">这次回答用的型号。</param>
        /// <param name="apiKey">已经解开的密钥。</param>
        private async void BeginTitle(int serial, AiConversation conversation, ModelPick model, string apiKey)
        {
            try
            {
                if (conversation == null || model.Profile == null) return;
                string prompt = AiPrompt.Title(conversation.Title, AiAssistant.RecentVisible(conversation.Messages, 1500));
                AiAuditEntry entry = new AiAuditEntry
                {
                    Action = "chat.title",
                    ConversationId = conversation.Id,
                    ConversationTitle = conversation.Title,
                    ModelDisplayName = AiModelCaption.Audit(model.Profile.DisplayName, model.ModelId),
                    Protocol = model.Profile.Protocol,
                    Endpoint = AiEndpoints.Describe(model.Profile.Endpoint),
                    UserText = prompt,
                    HasTab = false,
                    Redaction = "无"
                };
                AiReply reply = await _assistant.AskAsync(
                    model.Profile,
                    apiKey,
                    model.ModelId,
                    entry,
                    null,
                    null,
                    prompt);
                if (serial != _titleSerial) return;
                RunOnUiThread(() =>
                {
                    if (IsDisposed || serial != _titleSerial) return;
                    _assistant.Conversations.SetGeneratedTitle(conversation.Id, reply == null ? null : reply.Text);
                });
            }
            catch (Exception)
            {
                // 标题是附带请求。主回复已经成功，失败或审计写失败时保留原标题，不弹错误详情。
            }
        }

        /// <summary>未锁定，且这一轮是第 1、3、7 或 13 次用户消息时才刷新标题。</summary>
        /// <param name="conversation">刚写入回复的对话。</param>
        /// <returns>需要另起标题请求时为 true。第 13 轮之后不再自动改。</returns>
        private static bool ShouldRefreshTitle(AiConversation conversation)
        {
            if (conversation == null || conversation.TitleLocked || conversation.Messages == null) return false;
            int users = 0;
            foreach (AiChatTurn turn in conversation.Messages)
            {
                if (turn != null && string.Equals(turn.Role, "user", StringComparison.OrdinalIgnoreCase))
                    users++;
            }

            return users == 1 || users == 3 || users == 7 || users == 13;
        }

        private void ShowTranscript()
        {
            if (!_browserReady || _view.CoreWebView2 == null) return;
            AiConversation conversation = Current();
            bool pendingHere = !string.IsNullOrEmpty(_pendingUser)
                && string.Equals(_currentId, _pendingConversationId, StringComparison.OrdinalIgnoreCase);
            AiChatTurn pending = null;
            if (pendingHere)
            {
                pending = new AiChatTurn
                {
                    Role = "user",
                    Text = _pendingUser,
                    Sent = _pendingSent
                };
            }

            string html = AiTranscriptHtml.Document(
                conversation == null ? null : conversation.Messages,
                pending,
                pendingHere && _sending);
            _view.CoreWebView2.NavigateToString(html);
        }

        private bool Select(string id)
        {
            for (int index = 0; index < _tabs.TabPages.Count; index++)
            {
                if (!string.Equals(_tabs.TabPages[index].Tag as string, id, StringComparison.OrdinalIgnoreCase))
                    continue;
                _binding = true;
                _tabs.SelectedIndex = index;
                _currentId = id;
                _binding = false;
                ClearConversationAlerts(_currentId);
                RefreshConversationTabIcons();
                ShowTranscript();
                return true;
            }

            return false;
        }

        /// <summary>把会话标记成失败，并清掉该会话的未读标记。</summary>
        /// <param name="conversationId">会话标识。</param>
        private void MarkConversationError(string conversationId)
        {
            if (string.IsNullOrEmpty(conversationId)) return;
            _unreadConversationIds.Remove(conversationId);
            _errorConversationIds.Add(conversationId);
        }

        /// <summary>把会话标记成未读，并清掉失败标记。</summary>
        /// <param name="conversationId">会话标识。</param>
        private void MarkConversationUnread(string conversationId)
        {
            if (string.IsNullOrEmpty(conversationId)) return;
            _errorConversationIds.Remove(conversationId);
            _unreadConversationIds.Add(conversationId);
        }

        /// <summary>用户查看会话后清除失败和未读提示。</summary>
        /// <param name="conversationId">会话标识。</param>
        private void ClearConversationAlerts(string conversationId)
        {
            if (string.IsNullOrEmpty(conversationId)) return;
            _errorConversationIds.Remove(conversationId);
            _unreadConversationIds.Remove(conversationId);
        }

        private void SelectFirstActive()
        {
            if (_tabs.TabPages.Count == 0)
            {
                _currentId = null;
                ShowTranscript();
                return;
            }

            Select(_tabs.TabPages[0].Tag as string);
        }

        private bool TabExists(string id)
        {
            for (int index = 0; index < _tabs.TabPages.Count; index++)
            {
                if (string.Equals(_tabs.TabPages[index].Tag as string, id, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private int TabIndexAt(Point location)
        {
            for (int index = 0; index < _tabs.TabCount; index++)
            {
                if (_tabs.GetTabRect(index).Contains(location)) return index;
            }

            return -1;
        }

        /// <summary>第一次装入列表时读出上次勾选。之后以内存为准，避免盖掉刚改的选择。</summary>
        private void LoadSkillChoice()
        {
            if (_skillChoiceLoaded) return;
            _skillChoiceLoaded = true;
            string stored = _settings.Get(CoreSettings.AiSkillChoice) ?? string.Empty;
            _skillIds.Clear();
            if (stored.Length == 0) return;
            foreach (string part in stored.Split('|'))
            {
                string id = part.Trim();
                if (id.Length == 0 || SkillChosen(id)) continue;
                _skillIds.Add(id);
            }
        }

        /// <summary>把当前勾选写成竖线分隔的标识。一条都没勾时写成空字符串。</summary>
        private void RememberSkillChoice()
        {
            string text = string.Join("|", _skillIds.ToArray());
            string stored = _settings.Get(CoreSettings.AiSkillChoice) ?? string.Empty;
            if (string.Equals(stored, text, StringComparison.Ordinal)) return;
            _settings.Set(CoreSettings.AiSkillChoice, text);
        }

        /// <summary>去掉已经不在列表里的勾选。剩下的仍按菜单顺序发送。变少时写回会话。</summary>
        private void PruneSkillIds()
        {
            string before = string.Join("|", _skillIds.ToArray());
            List<string> kept = new List<string>();
            foreach (SkillPick pick in _skills)
            {
                if (pick.Skill == null || string.IsNullOrEmpty(pick.Skill.Id)) continue;
                if (SkillChosen(pick.Skill.Id))
                    kept.Add(pick.Skill.Id);
            }

            _skillIds.Clear();
            _skillIds.AddRange(kept);
            if (!string.Equals(before, string.Join("|", _skillIds.ToArray()), StringComparison.Ordinal))
                RememberSkillChoice();
        }

        /// <summary>这条 Skill 是否在本次发送里。</summary>
        /// <param name="id">Skill 标识。</param>
        /// <returns>已勾选时为 true。</returns>
        private bool SkillChosen(string id)
        {
            foreach (string existing in _skillIds)
            {
                if (string.Equals(existing, id, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        /// <summary>按菜单顺序取出已勾选的 Skill。「不使用」不算。</summary>
        /// <returns>本次要附带的 Skill。</returns>
        private List<SkillPick> SelectedSkills()
        {
            List<SkillPick> chosen = new List<SkillPick>();
            foreach (SkillPick pick in _skills)
            {
                if (pick.Skill == null || !SkillChosen(pick.Skill.Id)) continue;
                chosen.Add(pick);
            }

            return chosen;
        }

        private void FillChoiceMenu(ContextMenuStrip menu, bool model)
        {
            menu.Items.Clear();
            if (model)
            {
                foreach (ModelPick pick in _models)
                {
                    ModelPick captured = pick;
                    ToolStripMenuItem item = new ToolStripMenuItem(captured.MenuText, AiModelIcons.For(captured.Profile.DisplayName, new[] { captured.ModelId }), (sender, args) =>
                    {
                        _modelId = captured.Profile.Id;
                        _modelName = captured.ModelId;
                        RememberModel();
                        ApplyChoiceButtons();
                    });
                    item.Checked = SameModel(captured);
                    menu.Items.Add(item);
                }

                AppendSettingsItem(menu, AppPresentation.AiAddModelMenu, OpenModelSettings);
                return;
            }

            foreach (SkillPick pick in _skills)
            {
                bool clear = pick.Skill == null;
                string id = clear ? string.Empty : pick.Skill.Id ?? string.Empty;
                ToolStripMenuItem item = new ToolStripMenuItem(pick.ToString(), null, (sender, args) => ToggleSkill(clear, id));
                item.Checked = !clear && SkillChosen(id);
                menu.Items.Add(item);
            }

            AppendSettingsItem(menu, AppPresentation.AiAddSkillMenu, OpenSkillSettings);
        }

        /// <summary>在选择菜单末尾加上打开设置的一项。有上方条目时先加分隔线。</summary>
        /// <param name="menu">模型或 Skill 菜单。调用前已经写完可选项。</param>
        /// <param name="text">末项文字。</param>
        /// <param name="open">点该项后要打开的设置窗。在界面线程上执行。</param>
        /// <remarks>菜单关闭后再打开模态窗，避免下拉还没收起就把设置窗压在后面。不改当前选择。</remarks>
        private void AppendSettingsItem(ContextMenuStrip menu, string text, Action open)
        {
            if (menu.Items.Count > 0)
                menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(new ToolStripMenuItem(text, _addIcon, (sender, args) =>
            {
                if (IsDisposed || !IsHandleCreated) return;
                BeginInvoke(new Action(() =>
                {
                    if (IsDisposed) return;
                    Guard(open);
                }));
            }));
        }

        /// <summary>以模态方式打开模型设置。已有配置时仍选中第一项。</summary>
        /// <remarks>必须在界面线程调用。所有者是查询窗或报错对话窗。关闭后由模型存储的变更事件刷新菜单。</remarks>
        private void OpenModelSettings()
        {
            using (AiModelSettingsForm form = new AiModelSettingsForm(_assistant, _settings))
                form.ShowDialog(FindForm());
        }

        /// <summary>以模态方式打开 Skill 设置。已有 Skill 时仍选中第一项。</summary>
        /// <remarks>必须在界面线程调用。所有者是查询窗或报错对话窗。关闭后由 Skill 存储的变更事件刷新菜单。</remarks>
        private void OpenSkillSettings()
        {
            using (AiSkillSettingsForm form = new AiSkillSettingsForm(_assistant, _settings))
                form.ShowDialog(FindForm());
        }

        private ModelPick CurrentModel()
        {
            foreach (ModelPick pick in _models)
            {
                if (SameModel(pick))
                    return pick;
            }

            return null;
        }

        private bool SameModel(ModelPick pick)
        {
            return pick != null
                && string.Equals(pick.Profile.Id, _modelId, StringComparison.OrdinalIgnoreCase)
                && string.Equals(pick.ModelId, _modelName, StringComparison.OrdinalIgnoreCase);
        }

        private void RememberModel()
        {
            if (string.IsNullOrEmpty(_modelId) || string.IsNullOrEmpty(_modelName)) return;
            _settings.Set(CoreSettings.AiModelChoice, _modelId + "|" + _modelName);
        }

        /// <summary>勾选或取消一条 Skill。点「不使用」则清掉全部并让菜单收起。</summary>
        /// <param name="clear">为 true 时清掉全部勾选。</param>
        /// <param name="id">要切换的 Skill 标识。清掉时忽略。</param>
        private void ToggleSkill(bool clear, string id)
        {
            if (clear)
            {
                _keepSkillMenu = false;
                _skillIds.Clear();
            }
            else
            {
                _keepSkillMenu = true;
                if (SkillChosen(id))
                    _skillIds.RemoveAll(existing => string.Equals(existing, id, StringComparison.OrdinalIgnoreCase));
                else if (!string.IsNullOrEmpty(id))
                    _skillIds.Add(id);
            }

            ApplyChoiceButtons();
            SyncSkillChecks();
            RememberSkillChoice();
        }

        /// <summary>勾选 Skill 时不收起菜单，方便连续勾。点「不使用」或末项时照常收起。</summary>
        /// <param name="sender">Skill 菜单。</param>
        /// <param name="e">关闭原因。取消后菜单留在原处。</param>
        private void SkillMenuClosing(object sender, ToolStripDropDownClosingEventArgs e)
        {
            if (e.CloseReason == ToolStripDropDownCloseReason.ItemClicked && _keepSkillMenu)
                e.Cancel = true;
            _keepSkillMenu = false;
        }

        /// <summary>按当前勾选刷新菜单里的对勾，不重建菜单。</summary>
        private void SyncSkillChecks()
        {
            int index = 0;
            foreach (SkillPick pick in _skills)
            {
                if (index >= _skillMenu.Items.Count) return;
                ToolStripMenuItem item = _skillMenu.Items[index] as ToolStripMenuItem;
                if (item == null) return;
                string id = pick.Skill == null ? string.Empty : pick.Skill.Id ?? string.Empty;
                item.Checked = pick.Skill != null && SkillChosen(id);
                index++;
            }
        }

        private string CurrentModelText()
        {
            ModelPick model = CurrentModel();
            return model == null ? "未设置模型" : model.ModelId;
        }

        /// <summary>芯片上的 Skill 文字。没选是 Skills，一条是名称，多条是个数。</summary>
        /// <returns>不含下拉标记的文字。</returns>
        private string CurrentSkillText()
        {
            List<SkillPick> chosen = SelectedSkills();
            if (chosen.Count == 0) return "Skills";
            if (chosen.Count == 1) return chosen[0].Skill.Name;
            return chosen.Count.ToString() + " 个 Skill";
        }

        private void ApplyChoiceButtons()
        {
            ModelPick model = CurrentModel();
            _modelButton.Text = CurrentModelText() + "  ﹀";
            _modelButton.Image = model == null
                ? AiModelIcons.For(string.Empty, null)
                : AiModelIcons.For(model.Profile.DisplayName, new[] { model.ModelId });
            bool skillOn = SelectedSkills().Count > 0;
            _skillButton.Text = CurrentSkillText() + "  ﹀";
            _skillButton.Image = AiModelIcons.Skill(skillOn);
            LayoutComposer();
        }

        private AiConversation Current()
        {
            if (string.IsNullOrEmpty(_currentId)) return null;
            return _assistant.Conversations.Get(_currentId);
        }

        private void RefreshSqlReference()
        {
            if (_omitSql)
            {
                _sqlRef.Visible = false;
                return;
            }

            AiEditorSnapshot snapshot = _readEditor() ?? new AiEditorSnapshot();
            bool show = snapshot.HasTab && !string.IsNullOrWhiteSpace(snapshot.TabName);
            _sqlRef.Visible = show;
            _sqlRef.Text = show ? CloseMark.Caption(snapshot.TabName, _sqlRef.Font) : string.Empty;
        }

        private void RunOnUiThread(Action action)
        {
            if (action == null) return;
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                if (!IsHandleCreated) return;
                Invoke(action);
                return;
            }

            action();
        }

        private static string CombineNotes(string left, string right)
        {
            if (left == "无") return right ?? "无";
            if (right == "无" || string.IsNullOrEmpty(right) || left == right) return left;
            return left + "、" + right;
        }

        /// <summary>引用标签。文字用链接色，左侧是类型图标，左右留出间隔。</summary>
        /// <param name="kind">查询方案、预览、文件或其他。</param>
        /// <returns>尚未写入文字的平按钮。调用方设置文字并接上关闭。</returns>
        private Button ReferenceChip(MentionKind kind)
        {
            Color chip = ChipColor();
            Button button = CloseMark.CreateChip();
            button.AutoSize = true;
            button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            button.FlatStyle = FlatStyle.Flat;
            button.BackColor = chip;
            button.ForeColor = SystemColors.HotTrack;
            button.Font = _chipFont;
            button.Image = MentionIcon(kind, true);
            button.TextImageRelation = TextImageRelation.ImageBeforeText;
            button.ImageAlign = ContentAlignment.MiddleLeft;
            button.TextAlign = ContentAlignment.MiddleLeft;
            button.Margin = new Padding(4, 2, 4, 0);
            button.Padding = new Padding(2, 0, 2, 0);
            button.Cursor = Cursors.Hand;
            button.UseVisualStyleBackColor = false;
            button.TabStop = false;
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = SystemColors.ControlDark;
            button.FlatAppearance.MouseDownBackColor = SystemColors.ControlDark;
            return button;
        }

        /// <summary>按引用类型取 16 像素图标。不可用项用灰色，避免和能点的行混在一起。</summary>
        /// <param name="kind">查询方案、预览、文件或其他。</param>
        /// <param name="enabled">为 false 时返回灰度图。</param>
        /// <returns>本控件缓存的图像。不要释放。</returns>
        private Image MentionIcon(MentionKind kind, bool enabled)
        {
            string fileName;
            switch (kind)
            {
                case MentionKind.Query:
                    fileName = "icon_script.png";
                    break;
                case MentionKind.Preview:
                    fileName = "icon_table.png";
                    break;
                case MentionKind.File:
                    fileName = "icon_file.png";
                    break;
                default:
                    fileName = "icon_attachment.png";
                    break;
            }

            return CachedMentionIcon(fileName, !enabled);
        }

        /// <summary>加载并记住引用图标。同名同色只加载一次。</summary>
        /// <param name="fileName">Resources 下的文件名。</param>
        /// <param name="grayscale">为 true 时要灰度图。</param>
        /// <returns>缓存中的图像。</returns>
        private Image CachedMentionIcon(string fileName, bool grayscale)
        {
            string key = (grayscale ? "gray:" : "color:") + (fileName ?? string.Empty);
            Image image;
            if (_mentionIcons.TryGetValue(key, out image))
                return image;
            image = DatabaseTreeImageCatalog.LoadActionIcon(fileName, 16, grayscale);
            _mentionIcons.Add(key, image);
            return image;
        }

        private static FlowLayoutPanel ChoiceRow()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                WrapContents = false,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(8, 4, 8, 0),
                BackColor = SystemColors.Control
            };
        }

        /// <summary>芯片比窗体底色略深，跟着系统主题走，不写死 RGB。</summary>
        private static Color ChipColor()
        {
            Color basis = SystemColors.Control;
            Color dark = SystemColors.ControlDark;
            return Color.FromArgb(
                (basis.R * 3 + dark.R) / 4,
                (basis.G * 3 + dark.G) / 4,
                (basis.B * 3 + dark.B) / 4);
        }

        /// <summary>Skill 和模型用的平芯片。字号与灰色说明相同。</summary>
        private Button QuietButton()
        {
            Color chip = ChipColor();
            Button button = new Button
            {
                AutoSize = true,
                Font = _chipFont,
                FlatStyle = FlatStyle.Flat,
                BackColor = chip,
                Margin = new Padding(0, 0, 4, 0),
                UseVisualStyleBackColor = false
            };
            button.FlatAppearance.BorderSize = 0;
            button.FlatAppearance.MouseOverBackColor = SystemColors.ControlDark;
            button.FlatAppearance.MouseDownBackColor = SystemColors.ControlDarkDark;
            return button;
        }

        private sealed class ModelPick
        {
            public ModelPick(AiModelProfile profile, string modelId)
            {
                Profile = profile;
                ModelId = modelId ?? string.Empty;
            }

            public AiModelProfile Profile { get; }
            public string ModelId { get; }
            public string MenuText
            {
                get { return AiModelCaption.Menu(Profile == null ? string.Empty : Profile.DisplayName, ModelId); }
            }
        }

        private sealed class SkillPick
        {
            public SkillPick(AiSkill skill) { Skill = skill; }
            public AiSkill Skill { get; }
            public override string ToString() { return Skill == null ? "不使用" : Skill.Name; }
        }

        private sealed class Attachment
        {
            public string Label { get; set; }
            public string Body { get; set; }
            public MentionKind Kind { get; set; }
            public bool Preview { get { return Kind == MentionKind.Preview; } }
            public Control View { get; set; }
        }

        private enum MentionKind
        {
            File,
            Preview,
            Query
        }

        private sealed class MentionChoice
        {
            public MentionChoice(string text, MentionKind kind, string label, string body)
            {
                Text = text ?? string.Empty;
                Kind = kind;
                Label = label ?? string.Empty;
                Body = body ?? string.Empty;
                Enabled = true;
            }

            public string Text { get; }
            public MentionKind Kind { get; }
            public string Label { get; }
            public string Body { get; }
            public bool Enabled { get; set; }
        }

        /// <summary>停在输入框旁的引用菜单。不抢走输入框焦点，方便继续用键盘筛选。</summary>
        private sealed class MentionPopup : Form
        {
            private readonly ListBox _list = new ListBox
            {
                Dock = DockStyle.Fill,
                IntegralHeight = false,
                BorderStyle = BorderStyle.None
            };
            private readonly Action<MentionChoice> _accept;
            private readonly List<MentionChoice> _choices = new List<MentionChoice>();

            private readonly Func<MentionKind, bool, Image> _iconOf;

            public MentionPopup(Action<MentionChoice> accept, Func<MentionKind, bool, Image> iconOf)
            {
                _accept = accept ?? throw new ArgumentNullException(nameof(accept));
                _iconOf = iconOf;
                FormBorderStyle = FormBorderStyle.FixedSingle;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                ControlBox = false;
                Width = 280;
                Height = 220;
                Controls.Add(_list);
                _list.DrawMode = DrawMode.OwnerDrawFixed;
                _list.ItemHeight = 24;
                _list.DrawItem += DrawChoice;
                _list.Click += (sender, args) => AcceptCurrent();
            }

            private void DrawChoice(object sender, DrawItemEventArgs e)
            {
                if (e.Index < 0 || e.Index >= _choices.Count) return;
                MentionChoice choice = _choices[e.Index];
                bool selected = choice.Enabled && (e.State & DrawItemState.Selected) != 0;
                e.Graphics.FillRectangle(selected ? SystemBrushes.Highlight : SystemBrushes.Window, e.Bounds);
                Image icon = _iconOf == null ? null : _iconOf(choice.Kind, choice.Enabled);
                Rectangle text = e.Bounds;
                if (icon != null)
                {
                    int y = e.Bounds.Top + Math.Max(0, (e.Bounds.Height - icon.Height) / 2);
                    e.Graphics.DrawImage(icon, e.Bounds.Left + 4, y, icon.Width, icon.Height);
                    text = new Rectangle(e.Bounds.Left + icon.Width + 8, e.Bounds.Top, Math.Max(0, e.Bounds.Width - icon.Width - 8), e.Bounds.Height);
                }

                TextRenderer.DrawText(
                    e.Graphics,
                    choice.Text,
                    e.Font,
                    text,
                    !choice.Enabled ? SystemColors.GrayText : (selected ? SystemColors.HighlightText : SystemColors.ControlText),
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            }

            protected override bool ShowWithoutActivation
            {
                get { return true; }
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams parameters = base.CreateParams;
                    parameters.ExStyle |= 0x08000000;
                    return parameters;
                }
            }

            public void ShowChoices(IList<MentionChoice> choices, Point screen, Form owner)
            {
                _choices.Clear();
                _list.Items.Clear();
                if (choices != null)
                {
                    foreach (MentionChoice choice in choices)
                    {
                        _choices.Add(choice);
                        _list.Items.Add(choice.Text);
                    }
                }

                int index = FirstEnabled();
                if (index >= 0)
                    _list.SelectedIndex = index;
                Location = screen;
                if (!Visible)
                {
                    if (owner != null && !owner.IsDisposed)
                        Show(owner);
                    else
                        Show();
                }
            }

            public void MoveSelection(int delta)
            {
                if (_choices.Count == 0 || delta == 0) return;
                int index = _list.SelectedIndex;
                if (index < 0)
                    index = delta > 0 ? -1 : _choices.Count;
                for (int step = 0; step < _choices.Count; step++)
                {
                    index += delta;
                    if (index < 0)
                        index = _choices.Count - 1;
                    if (index >= _choices.Count)
                        index = 0;
                    if (_choices[index].Enabled)
                    {
                        _list.SelectedIndex = index;
                        return;
                    }
                }
            }

            public void AcceptCurrent()
            {
                int index = _list.SelectedIndex;
                if (index < 0 || index >= _choices.Count || !_choices[index].Enabled) return;
                MentionChoice choice = _choices[index];
                Hide();
                _accept(choice);
            }

            private int FirstEnabled()
            {
                for (int index = 0; index < _choices.Count; index++)
                {
                    if (_choices[index].Enabled)
                        return index;
                }

                return -1;
            }
        }
    }

    /// <summary>为本进程准备一份 WebView2 环境。加载器按 Excel 位数从插件目录读取。</summary>
    /// <remarks>用户数据目录由调用方传入。多个对话控件共用这一份环境。</remarks>
    internal static class AiChatBrowserHost
    {
        private static readonly object Gate = new object();
        private static Task<CoreWebView2Environment> _environment;
        private static bool _loaderReady;

        /// <summary>在创建任何 WebView2 控件之前指定加载器目录。同一进程只调用一次。</summary>
        /// <remarks>关闭查询窗后环境仍留在进程里。再次调用 SetLoaderDllFolderPath 会抛出 InvalidOperationException。</remarks>
        public static void PrepareLoader()
        {
            if (_loaderReady) return;
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string bitness = IntPtr.Size == 8 ? "x64" : "x86";
            string runtimes = Path.Combine(baseDir, "runtimes", "win-" + bitness, "native");
            string linked = Path.Combine(baseDir, bitness);
            string folder = Directory.Exists(runtimes) ? runtimes : linked;
            CoreWebView2Environment.SetLoaderDllFolderPath(folder);
            _loaderReady = true;
        }

        /// <summary>创建或复用浏览器环境。</summary>
        /// <param name="userData">可写的用户数据目录。</param>
        /// <returns>环境任务。失败时任务为失败状态，调用方显示缺少运行库。</returns>
        public static Task<CoreWebView2Environment> CreateAsync(string userData)
        {
            lock (Gate)
            {
                if (_environment != null) return _environment;
                Directory.CreateDirectory(userData ?? string.Empty);
                _environment = CoreWebView2Environment.CreateAsync(null, userData);
                return _environment;
            }
        }
    }
}
