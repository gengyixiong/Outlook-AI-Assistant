using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Diagnostics;
using OutlookAiAssistant.Models;
using OutlookAiAssistant.OutlookIntegration;
using OutlookAiAssistant.SemanticSearch;

namespace OutlookAiAssistant.UI
{
    /// <summary>
    /// The COM-visible control hosted in Outlook's right-side custom task pane.
    /// All Outlook and AI actions are user initiated from this control.
    /// </summary>
    [ComVisible(true)]
    [Guid("60DF60D1-BB82-4606-B245-F79DD62C99AC")]
    [ProgId("OutlookAiAssistant.TaskPane")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AssistantPaneControl : UserControl
    {
        private TabControl _tabs;
        private Button _summarizeButton;
        private Label _mailInfo;
        private RichTextBox _summaryOutput;
        private Button _copySummaryButton;
        private Label _status;
        private Label _configurationStatus;
        private bool _busy;

        // Semantic search page.
        private TextBox _semanticInput;
        private Button _semanticSearchButton;
        private Button _semanticIndexButton;
        private Button _semanticStopButton;
        private Button _semanticDownloadModelButton;
        private Button _semanticStartServiceButton;
        private Button _semanticRestartServiceButton;
        private Button _semanticStopServiceButton;
        private Button _openServiceButton;
        private Label _semanticServiceStatus;
        private Label _semanticIndexStatus;
        private TableLayoutPanel _semanticResultsPanel;
        private Label _semanticStatus;
        private bool _semanticBusy;
        private CancellationTokenSource _indexCancellation;

        // In-place Settings page.
        private TextBox _settingsBaseUrl;
        private TextBox _settingsApiKey;
        private CheckBox _settingsShowApiKey;
        private CheckBox _settingsClearApiKey;
        private ComboBox _settingsModel;
        private ComboBox _settingsReasoningEffort;
        private ComboBox _settingsLanguage;
        private NumericUpDown _settingsMaxCharacters;
        private TextBox _settingsIdentityEmails;
        private TextBox _settingsIdentityAliases;
        private Button _settingsSaveButton;
        private Label _settingsSaveStatus;
        private Button _openModelDirButton;
        private Button _openIndexDirButton;
        private Button _openLogDirButton;

        public AssistantPaneControl()
        {
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            BuildInterface();
            Load += delegate
            {
                LoadSettingsToUI();
                RefreshSemanticIndexStatus();
            };
        }

        private void BuildInterface()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Fill;
            _tabs.Margin = new Padding(10, 8, 10, 4);
            _tabs.TabPages.Add(BuildSummaryPage());
            _tabs.TabPages.Add(BuildSemanticSearchPage());
            _tabs.TabPages.Add(BuildSettingsPage());
            root.Controls.Add(_tabs);

            _status = new Label();
            _status.Text = "就绪";
            _status.AutoEllipsis = true;
            _status.Dock = DockStyle.Bottom;
            _status.Padding = new Padding(12, 7, 12, 7);
            _status.BackColor = Color.FromArgb(245, 245, 245);
            root.Controls.Add(_status);
        }

        private TabPage BuildSummaryPage()
        {
            TabPage page = new TabPage("邮件摘要");
            page.Padding = new Padding(10);
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            _mailInfo = new Label();
            _mailInfo.Text = "尚未读取邮件。";
            _mailInfo.AutoSize = true;
            _mailInfo.MaximumSize = new Size(350, 0);
            _mailInfo.Margin = new Padding(3, 3, 3, 10);
            layout.Controls.Add(_mailInfo);

            _summarizeButton = CreatePrimaryButton("生成 Executive Brief");
            _summarizeButton.Click += SummarizeClicked;
            layout.Controls.Add(_summarizeButton);

            _summaryOutput = new RichTextBox();
            _summaryOutput.Dock = DockStyle.Fill;
            _summaryOutput.ReadOnly = true;
            _summaryOutput.BackColor = Color.White;
            _summaryOutput.BorderStyle = BorderStyle.FixedSingle;
            _summaryOutput.Margin = new Padding(3, 10, 3, 6);
            layout.Controls.Add(_summaryOutput);

            _copySummaryButton = new Button();
            _copySummaryButton.Text = "复制摘要";
            _copySummaryButton.AutoSize = true;
            _copySummaryButton.Enabled = false;
            _copySummaryButton.Click += delegate
            {
                if (!string.IsNullOrWhiteSpace(_summaryOutput.Text))
                {
                    Clipboard.SetText(_summaryOutput.Text);
                    SetStatus("摘要已复制到剪贴板。");
                }
            };
            layout.Controls.Add(_copySummaryButton);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildSemanticSearchPage()
        {
            TabPage page = new TabPage("语义搜索");
            page.Padding = new Padding(10);
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 7;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Service + model + index status, always visible at the top.
            _semanticServiceStatus = new Label();
            _semanticServiceStatus.AutoSize = true;
            _semanticServiceStatus.Dock = DockStyle.Top;
            _semanticServiceStatus.Font = new Font(Font, FontStyle.Bold);
            _semanticServiceStatus.Text = "服务状态：正在检查…";
            layout.Controls.Add(_semanticServiceStatus);

            _semanticInput = new TextBox();
            _semanticInput.Multiline = true;
            _semanticInput.ScrollBars = ScrollBars.Vertical;
            _semanticInput.Height = 70;
            _semanticInput.Dock = DockStyle.Top;
            _semanticInput.Margin = new Padding(3, 8, 3, 8);
            layout.Controls.Add(_semanticInput);

            _semanticSearchButton = CreatePrimaryButton("语义搜索");
            _semanticSearchButton.Click += SemanticSearchClicked;
            layout.Controls.Add(_semanticSearchButton);

            FlowLayoutPanel actions = new FlowLayoutPanel();
            actions.AutoSize = true;
            actions.Dock = DockStyle.Top;
            actions.WrapContents = true;
            actions.Margin = new Padding(0, 8, 0, 0);

            _semanticIndexButton = new Button();
            _semanticIndexButton.Text = "建立 / 更新语义索引";
            _semanticIndexButton.AutoSize = true;
            _semanticIndexButton.MinimumSize = new Size(150, 29);
            _semanticIndexButton.Margin = new Padding(3, 0, 3, 3);
            _semanticIndexButton.Click += SemanticIndexClicked;
            actions.Controls.Add(_semanticIndexButton);

            _semanticStopButton = new Button();
            _semanticStopButton.Text = "停止索引";
            _semanticStopButton.AutoSize = true;
            _semanticStopButton.MinimumSize = new Size(88, 29);
            _semanticStopButton.Enabled = false;
            _semanticStopButton.Margin = new Padding(3, 0, 3, 3);
            _semanticStopButton.Click += SemanticStopClicked;
            actions.Controls.Add(_semanticStopButton);

            _semanticStartServiceButton = new Button();
            _semanticStartServiceButton.Text = "启动服务";
            _semanticStartServiceButton.AutoSize = true;
            _semanticStartServiceButton.MinimumSize = new Size(88, 29);
            _semanticStartServiceButton.Margin = new Padding(3, 0, 3, 3);
            _semanticStartServiceButton.Click += SemanticStartServiceClicked;
            actions.Controls.Add(_semanticStartServiceButton);

            _semanticRestartServiceButton = new Button();
            _semanticRestartServiceButton.Text = "重启服务";
            _semanticRestartServiceButton.AutoSize = true;
            _semanticRestartServiceButton.MinimumSize = new Size(88, 29);
            _semanticRestartServiceButton.Margin = new Padding(3, 0, 3, 3);
            _semanticRestartServiceButton.Click += SemanticRestartServiceClicked;
            actions.Controls.Add(_semanticRestartServiceButton);

            _semanticStopServiceButton = new Button();
            _semanticStopServiceButton.Text = "停止服务";
            _semanticStopServiceButton.AutoSize = true;
            _semanticStopServiceButton.MinimumSize = new Size(88, 29);
            _semanticStopServiceButton.Margin = new Padding(3, 0, 3, 3);
            _semanticStopServiceButton.Click += SemanticStopServiceClicked;
            actions.Controls.Add(_semanticStopServiceButton);

            _openServiceButton = new Button();
            _openServiceButton.Text = "刷新状态";
            _openServiceButton.AutoSize = true;
            _openServiceButton.MinimumSize = new Size(88, 29);
            _openServiceButton.Margin = new Padding(3, 0, 3, 3);
            _openServiceButton.Click += CheckServiceClicked;
            actions.Controls.Add(_openServiceButton);
            layout.Controls.Add(actions);

            _semanticIndexStatus = new Label();
            _semanticIndexStatus.AutoSize = true;
            _semanticIndexStatus.Dock = DockStyle.Top;
            _semanticIndexStatus.ForeColor = Color.DimGray;
            _semanticIndexStatus.Text = "索引状态：未检查";
            layout.Controls.Add(_semanticIndexStatus);

            _semanticStatus = new Label();
            _semanticStatus.AutoSize = true;
            _semanticStatus.Dock = DockStyle.Top;
            _semanticStatus.Text = string.Empty;
            layout.Controls.Add(_semanticStatus);

            // A TableLayoutPanel with one 100%-wide column lets each result
            // card stretch to the current pane width. A FlowLayoutPanel would
            // keep the cards at a fixed width and clip them when the pane is
            // narrowed, which is why the cards used to look cut off.
            _semanticResultsPanel = new TableLayoutPanel();
            _semanticResultsPanel.Dock = DockStyle.Fill;
            _semanticResultsPanel.ColumnCount = 1;
            _semanticResultsPanel.AutoScroll = true;
            _semanticResultsPanel.AutoSize = false;
            _semanticResultsPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100F));
            _semanticResultsPanel.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_semanticResultsPanel);

            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildSettingsPage()
        {
            TabPage page = new TabPage("设置");
            page.Padding = new Padding(0);

            Panel scroll = new Panel();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;
            scroll.Padding = new Padding(12);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Top;
            layout.AutoSize = true;
            layout.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            layout.ColumnCount = 1;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // Status label at the top
            _configurationStatus = new Label();
            _configurationStatus.AutoSize = true;
            _configurationStatus.Dock = DockStyle.Top;
            _configurationStatus.Font = new Font(Font, FontStyle.Bold);
            _configurationStatus.Margin = new Padding(0, 0, 0, 10);
            layout.Controls.Add(_configurationStatus);

            // --- API Configuration Section Header ---
            Label apiHeader = new Label();
            apiHeader.Text = "API 接口配置 (OpenAI-Compatible)";
            apiHeader.Font = new Font(Font, FontStyle.Bold);
            apiHeader.AutoSize = true;
            apiHeader.Dock = DockStyle.Top;
            apiHeader.Margin = new Padding(0, 4, 0, 4);
            layout.Controls.Add(apiHeader);

            // Base URL
            Label urlLabel = new Label();
            urlLabel.Text = "Base URL：";
            urlLabel.AutoSize = true;
            urlLabel.Dock = DockStyle.Top;
            urlLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(urlLabel);

            _settingsBaseUrl = new TextBox();
            _settingsBaseUrl.Dock = DockStyle.Top;
            _settingsBaseUrl.Margin = new Padding(0, 0, 0, 6);
            layout.Controls.Add(_settingsBaseUrl);

            // API Key
            Label keyLabel = new Label();
            keyLabel.Text = "API Key：";
            keyLabel.AutoSize = true;
            keyLabel.Dock = DockStyle.Top;
            keyLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(keyLabel);

            _settingsApiKey = new TextBox();
            _settingsApiKey.Dock = DockStyle.Top;
            _settingsApiKey.UseSystemPasswordChar = true;
            _settingsApiKey.Margin = new Padding(0, 0, 0, 4);
            layout.Controls.Add(_settingsApiKey);

            FlowLayoutPanel keyOptions = new FlowLayoutPanel();
            keyOptions.AutoSize = true;
            keyOptions.Dock = DockStyle.Top;
            keyOptions.Margin = new Padding(0, 0, 0, 4);

            _settingsShowApiKey = new CheckBox();
            _settingsShowApiKey.Text = "显示密钥";
            _settingsShowApiKey.AutoSize = true;
            _settingsShowApiKey.Margin = new Padding(0, 0, 10, 0);
            _settingsShowApiKey.CheckedChanged += delegate
            {
                _settingsApiKey.UseSystemPasswordChar = !_settingsShowApiKey.Checked;
            };
            keyOptions.Controls.Add(_settingsShowApiKey);

            _settingsClearApiKey = new CheckBox();
            _settingsClearApiKey.Text = "清除已保存的密钥";
            _settingsClearApiKey.AutoSize = true;
            keyOptions.Controls.Add(_settingsClearApiKey);
            layout.Controls.Add(keyOptions);

            Label keyHint = new Label();
            keyHint.Text = "留空表示继续使用已保存的密钥。密钥使用 Windows DPAPI 本地加密存储。更改服务地址后需重新输入密钥。";
            keyHint.ForeColor = Color.DimGray;
            keyHint.AutoSize = true;
            keyHint.Dock = DockStyle.Top;
            keyHint.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(keyHint);

            // Model
            Label modelLabel = new Label();
            modelLabel.Text = "Model (模型名称)：";
            modelLabel.AutoSize = true;
            modelLabel.Dock = DockStyle.Top;
            modelLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(modelLabel);

            _settingsModel = new ComboBox();
            _settingsModel.Dock = DockStyle.Top;
            _settingsModel.DropDownStyle = ComboBoxStyle.DropDown;
            _settingsModel.Items.AddRange(new object[] { "gpt-5.6-luna", "gpt-4o", "deepseek-chat", "qwen-plus" });
            _settingsModel.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(_settingsModel);

            // Reasoning Effort
            Label effortLabel = new Label();
            effortLabel.Text = "Reasoning Effort (推理力度)：";
            effortLabel.AutoSize = true;
            effortLabel.Dock = DockStyle.Top;
            effortLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(effortLabel);

            _settingsReasoningEffort = new ComboBox();
            _settingsReasoningEffort.Dock = DockStyle.Top;
            _settingsReasoningEffort.DropDownStyle = ComboBoxStyle.DropDownList;
            _settingsReasoningEffort.Items.AddRange(new object[] { "None", "Medium", "Max" });
            _settingsReasoningEffort.SelectedIndex = 0;
            _settingsReasoningEffort.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(_settingsReasoningEffort);

            // Summary Language
            Label langLabel = new Label();
            langLabel.Text = "摘要语言：";
            langLabel.AutoSize = true;
            langLabel.Dock = DockStyle.Top;
            langLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(langLabel);

            _settingsLanguage = new ComboBox();
            _settingsLanguage.Dock = DockStyle.Top;
            _settingsLanguage.DropDownStyle = ComboBoxStyle.DropDownList;
            _settingsLanguage.Items.AddRange(new object[] { "简体中文", "English", "跟随邮件语言" });
            _settingsLanguage.SelectedIndex = 0;
            _settingsLanguage.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(_settingsLanguage);

            // Max Characters
            Label maxCharsLabel = new Label();
            maxCharsLabel.Text = "正文字符上限：";
            maxCharsLabel.AutoSize = true;
            maxCharsLabel.Dock = DockStyle.Top;
            maxCharsLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(maxCharsLabel);

            _settingsMaxCharacters = new NumericUpDown();
            _settingsMaxCharacters.Dock = DockStyle.Top;
            _settingsMaxCharacters.Minimum = 5000;
            _settingsMaxCharacters.Maximum = 200000;
            _settingsMaxCharacters.Increment = 5000;
            _settingsMaxCharacters.ThousandsSeparator = true;
            _settingsMaxCharacters.Value = 30000;
            _settingsMaxCharacters.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(_settingsMaxCharacters);

            // Identity Emails
            Label emailsLabel = new Label();
            emailsLabel.Text = "我的邮箱地址（每行一个）：";
            emailsLabel.AutoSize = true;
            emailsLabel.Dock = DockStyle.Top;
            emailsLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(emailsLabel);

            _settingsIdentityEmails = new TextBox();
            _settingsIdentityEmails.Dock = DockStyle.Top;
            _settingsIdentityEmails.Multiline = true;
            _settingsIdentityEmails.Height = 50;
            _settingsIdentityEmails.ScrollBars = ScrollBars.Vertical;
            _settingsIdentityEmails.Margin = new Padding(0, 0, 0, 8);
            layout.Controls.Add(_settingsIdentityEmails);

            // Identity Aliases
            Label aliasesLabel = new Label();
            aliasesLabel.Text = "别人对我的称俗（每行一个）：";
            aliasesLabel.AutoSize = true;
            aliasesLabel.Dock = DockStyle.Top;
            aliasesLabel.Margin = new Padding(0, 4, 0, 2);
            layout.Controls.Add(aliasesLabel);

            _settingsIdentityAliases = new TextBox();
            _settingsIdentityAliases.Dock = DockStyle.Top;
            _settingsIdentityAliases.Multiline = true;
            _settingsIdentityAliases.Height = 50;
            _settingsIdentityAliases.ScrollBars = ScrollBars.Vertical;
            _settingsIdentityAliases.Margin = new Padding(0, 0, 0, 10);
            layout.Controls.Add(_settingsIdentityAliases);

            // Save Button + Save Status
            FlowLayoutPanel saveRow = new FlowLayoutPanel();
            saveRow.AutoSize = true;
            saveRow.Dock = DockStyle.Top;
            saveRow.Margin = new Padding(0, 0, 0, 14);

            _settingsSaveButton = CreatePrimaryButton("保存设置");
            _settingsSaveButton.AutoSize = true;
            _settingsSaveButton.MinimumSize = new Size(110, 32);
            _settingsSaveButton.Click += SaveSettingsClicked;
            saveRow.Controls.Add(_settingsSaveButton);

            _settingsSaveStatus = new Label();
            _settingsSaveStatus.AutoSize = true;
            _settingsSaveStatus.Anchor = AnchorStyles.Left;
            _settingsSaveStatus.Margin = new Padding(8, 8, 0, 0);
            saveRow.Controls.Add(_settingsSaveStatus);
            layout.Controls.Add(saveRow);

            // --- Semantic Model & Local Data Section Header ---
            Label semanticHeader = new Label();
            semanticHeader.Text = "本地语义搜索模型与数据";
            semanticHeader.Font = new Font(Font, FontStyle.Bold);
            semanticHeader.AutoSize = true;
            semanticHeader.Dock = DockStyle.Top;
            semanticHeader.Margin = new Padding(0, 8, 0, 6);
            layout.Controls.Add(semanticHeader);

            _semanticDownloadModelButton = new Button();
            _semanticDownloadModelButton.Text = "下载语义搜索模型（约 1.5 GB）";
            _semanticDownloadModelButton.AutoSize = true;
            _semanticDownloadModelButton.MinimumSize = new Size(200, 30);
            _semanticDownloadModelButton.Margin = new Padding(0, 0, 0, 8);
            _semanticDownloadModelButton.Click += SemanticDownloadModelClicked;
            layout.Controls.Add(_semanticDownloadModelButton);

            FlowLayoutPanel dirButtons = new FlowLayoutPanel();
            dirButtons.AutoSize = true;
            dirButtons.Dock = DockStyle.Top;
            dirButtons.Margin = new Padding(0, 0, 0, 10);

            _openModelDirButton = new Button();
            _openModelDirButton.Text = "打开模型目录";
            _openModelDirButton.AutoSize = true;
            _openModelDirButton.Click += delegate { OpenDirectorySafely(SemanticServiceLauncher.DefaultModelDirectory); };
            dirButtons.Controls.Add(_openModelDirButton);

            _openIndexDirButton = new Button();
            _openIndexDirButton.Text = "打开索引目录";
            _openIndexDirButton.AutoSize = true;
            _openIndexDirButton.Click += delegate { OpenDirectorySafely(SemanticServiceLauncher.DefaultIndexDirectory); };
            dirButtons.Controls.Add(_openIndexDirButton);

            _openLogDirButton = new Button();
            _openLogDirButton.Text = "打开日志目录";
            _openLogDirButton.AutoSize = true;
            _openLogDirButton.Click += delegate { OpenDirectorySafely(SemanticServiceLauncher.DefaultLogDirectory); };
            dirButtons.Controls.Add(_openLogDirButton);
            layout.Controls.Add(dirButtons);

            Label note = new Label();
            note.AutoSize = true;
            note.ForeColor = Color.DimGray;
            note.Dock = DockStyle.Top;
            note.Text =
                "说明：API Key 使用 Windows DPAPI 加密保存在当前 Windows 用户目录中；"
                + "不会写入项目、安装目录或日志。更换电脑后需要重新填写。\r\n\r\n"
                + "本地语义模型与向量完全在本地运行，正文与向量绝不会上传至任何云端。";
            layout.Controls.Add(note);

            scroll.Controls.Add(layout);
            page.Controls.Add(scroll);
            return page;
        }

        private async void SummarizeClicked(object sender, EventArgs e)
        {
            if (_busy || !EnsureConfigured())
            {
                return;
            }

            SetBusy(true, "正在读取当前邮件及同一会话历史……");
            try
            {
                Connect addIn = RequireAddIn();
                // This is the only place where the UI asks Outlook for mail data.
                ConversationSnapshot conversation =
                    addIn.OutlookContext.GetCurrentConversation(
                        OutlookContextService
                            .DefaultMaximumConversationMessages);
                EmailSnapshot email = conversation.CurrentEmail;
                _mailInfo.Text =
                    "当前邮件：" + DisplaySubject(email.Subject)
                    + "\r\n独立历史邮件："
                    + conversation.HistoryMessages.Count
                    + "；Outlook 会话项目："
                    + conversation.TotalConversationItems
                    + (conversation.HistoryWasTruncated
                        ? "（已按上限取样）"
                        : string.Empty)
                    + "\r\n当前正文字符："
                    + (email.Body ?? string.Empty).Length
                    + "；附件：" + email.AttachmentCount;
                SetStatus("正在发送会话历史和当前邮件并生成 Executive Brief……");
                string summary = await addIn.SummaryService.SummarizeAsync(
                    conversation,
                    addIn.Settings,
                    CancellationToken.None);
                _summaryOutput.Text = summary;
                _copySummaryButton.Enabled = true;
                SetStatus("Executive Brief 已生成。");
            }
            catch (Exception ex)
            {
                ShowError("生成摘要失败", ex);
            }
            finally
            {
                SetBusy(false, null);
            }
        }

        private async void SemanticSearchClicked(object sender, EventArgs e)
        {
            if (_semanticBusy)
            {
                return;
            }

            Connect addIn = Connect.Current;
            if (addIn == null)
            {
                MessageBox.Show(
                    this,
                    "加载项尚未完成初始化，请重启 Outlook 后再试。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string description = _semanticInput.Text.Trim();
            if (description.Length == 0)
            {
                MessageBox.Show(
                    this,
                    "请先输入你要查找的邮件描述。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            SetSemanticBusy(true, "正在使用本地语义模型检索…");
            try
            {
                SemanticSearchFacade facade = addIn.CreateSemanticFacade();
                SemanticSearchFacade.Outcome outcome =
                    await facade.SearchAsync(description, CancellationToken.None);
                RenderSemanticResults(facade, outcome);
            }
            catch (Exception ex)
            {
                ShowError("语义搜索失败", ex);
            }
            finally
            {
                SetSemanticBusy(false, null);
            }
        }

        private void RenderSemanticResults(
            SemanticSearchFacade facade,
            SemanticSearchFacade.Outcome outcome)
        {
            _semanticResultsPanel.SuspendLayout();
            try
            {
                _semanticResultsPanel.Controls.Clear();
                _semanticResultsPanel.RowStyles.Clear();
                _semanticResultsPanel.RowCount = 0;
                if (!outcome.Success)
                {
                    _semanticStatus.Text = outcome.Message;
                    _semanticStatus.ForeColor = Color.FromArgb(176, 0, 32);
                    AddSemanticFallbackControls(facade);
                    return;
                }

                _semanticStatus.ForeColor = Color.DimGray;
                _semanticStatus.Text = string.IsNullOrWhiteSpace(outcome.Message)
                    ? ("找到 " + outcome.Response.Results.Count + " 条结果。")
                    : outcome.Message;

                for (int index = 0;
                    index < outcome.Response.Results.Count;
                    index++)
                {
                    AddFullWidthRow(
                        BuildResultCard(
                            facade,
                            outcome.Response.Results[index],
                            index + 1));
                }
            }
            finally
            {
                _semanticResultsPanel.ResumeLayout();
            }
        }

        /// <summary>
        /// Adds a control as a new row that spans the full panel width.
        /// </summary>
        private void AddFullWidthRow(Control control)
        {
            int row = _semanticResultsPanel.RowCount;
            _semanticResultsPanel.RowCount = row + 1;
            _semanticResultsPanel.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));
            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(3, 4, 3, 4);
            _semanticResultsPanel.Controls.Add(control, 0, row);
        }

        private void AddSemanticFallbackControls(SemanticSearchFacade facade)
        {
            Button checkService = new Button();
            checkService.Text = "检查本地服务状态";
            checkService.AutoSize = true;
            checkService.Click += CheckServiceClicked;
            AddFullWidthRow(checkService);

            Button outlookSearch = new Button();
            outlookSearch.Text = "改用 Outlook 原生搜索";
            outlookSearch.AutoSize = true;
            outlookSearch.Click += delegate
            {
                SelectSearchTab();
            };
            AddFullWidthRow(outlookSearch);
        }

        private Control BuildResultCard(
            SemanticSearchFacade facade,
            SemanticSearchResult result,
            int rank)
        {
            SemanticSearchResult captured = result;
            return BuildResultCardControl(
                result,
                rank,
                Font,
                delegate
                {
                    string error = facade.OpenResult(captured);
                    if (!string.IsNullOrWhiteSpace(error))
                    {
                        SetStatus(error);
                        MessageBox.Show(
                            this,
                            error,
                            "AI 邮件助手",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                    }
                },
                delegate { SearchInOutlook(captured); });
        }

        /// <summary>
        /// Builds one result card. Kept static and free of Outlook access so
        /// the layout can be unit tested.
        ///
        /// The card is width-agnostic: the outer frame docks to fill its table
        /// row, and the inner labels auto-size and wrap to whatever width the
        /// pane currently has. This lets the cards grow and shrink with the
        /// Outlook side pane. Two earlier problems are guarded against: nested
        /// Dock=Top panels collapsing to a sliver, and a fixed pixel width that
        /// clipped content when the pane was narrow.
        /// </summary>
        public static Control BuildResultCardControl(
            SemanticSearchResult result,
            int rank,
            Font baseFont,
            Action openAction,
            Action outlookAction)
        {
            TableLayoutPanel card = new TableLayoutPanel();
            card.ColumnCount = 1;
            card.Dock = DockStyle.Top;
            card.AutoSize = true;
            card.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            card.Padding = new Padding(8);
            card.Margin = new Padding(0);
            card.BackColor = Color.White;
            card.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;
            card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            Label title = new Label();
            title.Text = rank + ". " + DisplaySubjectStatic(result.Subject);
            title.Font = new Font(baseFont, FontStyle.Bold);
            title.AutoSize = true;
            title.Dock = DockStyle.Top;
            title.Margin = new Padding(0, 0, 0, 2);
            card.Controls.Add(title);

            Label meta = new Label();
            meta.Text = FormatSender(result) + "\r\n"
                + FormatDate(result.ReceivedAt)
                + (string.IsNullOrWhiteSpace(result.FolderPath)
                    ? string.Empty
                    : "  ·  " + result.FolderPath);
            meta.AutoSize = true;
            meta.Dock = DockStyle.Top;
            meta.ForeColor = Color.DimGray;
            meta.Margin = new Padding(0, 0, 0, 4);
            card.Controls.Add(meta);

            if (!string.IsNullOrWhiteSpace(result.AttachmentNames))
            {
                Label attachments = new Label();
                attachments.Text = "附件：" + result.AttachmentNames;
                attachments.AutoSize = true;
                attachments.Dock = DockStyle.Top;
                attachments.ForeColor = Color.DimGray;
                attachments.Margin = new Padding(0, 0, 0, 4);
                card.Controls.Add(attachments);
            }

            Label snippet = new Label();
            snippet.Text = string.IsNullOrWhiteSpace(result.Snippet)
                ? "（无正文片段）"
                : result.Snippet;
            snippet.AutoSize = true;
            snippet.Dock = DockStyle.Top;
            snippet.Margin = new Padding(0, 0, 0, 2);
            card.Controls.Add(snippet);

            Label hint = new Label();
            hint.Text = "双击卡片打开这封邮件";
            hint.AutoSize = true;
            hint.Dock = DockStyle.Top;
            hint.ForeColor = Color.FromArgb(0, 103, 184);
            hint.Margin = new Padding(0);
            card.Controls.Add(hint);

            // A Panel with BorderStyle draws the card outline reliably and
            // docks to the row width so it resizes with the pane.
            Panel frame = new Panel();
            frame.BorderStyle = BorderStyle.FixedSingle;
            frame.Dock = DockStyle.Fill;
            frame.AutoSize = true;
            frame.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            frame.Margin = new Padding(0);
            frame.Padding = new Padding(1);
            frame.Cursor = Cursors.Hand;
            frame.Controls.Add(card);

            // Double-clicking anywhere on the card opens the real mail. The
            // card no longer carries per-result buttons; the whole card is the
            // click target.
            EventHandler openHandler = delegate { openAction(); };
            frame.DoubleClick += openHandler;
            card.DoubleClick += openHandler;
            foreach (Control child in card.Controls)
            {
                child.Cursor = Cursors.Hand;
                child.DoubleClick += openHandler;
            }

            // Right-click still offers the Outlook search fallback without
            // cluttering the card.
            ContextMenuStrip menu = new ContextMenuStrip();
            menu.Items.Add("打开邮件", null, delegate { openAction(); });
            menu.Items.Add("在 Outlook 中搜索此标题", null,
                delegate { outlookAction(); });
            frame.ContextMenuStrip = menu;
            card.ContextMenuStrip = menu;
            foreach (Control child in card.Controls)
            {
                child.ContextMenuStrip = menu;
            }

            return frame;
        }

        private static string DisplaySubjectStatic(string subject)
        {
            string value = string.IsNullOrWhiteSpace(subject)
                ? "（无主题）"
                : subject.Trim();
            return value.Length <= 80 ? value : value.Substring(0, 80) + "…";
        }

        private void SearchInOutlook(SemanticSearchResult result)
        {
            try
            {
                Connect addIn = RequireAddIn();
                string subject = (result.Subject ?? string.Empty)
                    .Replace("\"", " ")
                    .Trim();
                if (subject.Length == 0)
                {
                    SetStatus("该结果没有可用于 Outlook 搜索的标题。");
                    return;
                }

                addIn.OutlookSearch.Search(
                    "subject:\"" + subject + "\"",
                    "all_folders");
                SetStatus("已将该邮件标题提交给 Outlook 原生搜索。");
            }
            catch (Exception ex)
            {
                ShowError("Outlook 搜索失败", ex);
            }
        }

        private async void SemanticIndexClicked(object sender, EventArgs e)
        {
            if (_semanticBusy)
            {
                return;
            }

            Connect addIn = Connect.Current;
            if (addIn == null)
            {
                MessageBox.Show(
                    this,
                    "加载项尚未完成初始化，请重启 Outlook 后再试。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            DialogResult confirm = MessageBox.Show(
                this,
                "将扫描最近 " + addIn.Settings.SemanticIndexMonths
                + " 个月、最多 " + addIn.Settings.SemanticIndexMaxMails
                + " 封邮件，读取正文并生成完全在本机运行的语义向量。"
                + "\r\n\r\n邮件正文和向量不会离开这台电脑。可以随时点击"
                + "“停止索引”中断。"
                + "\r\n\r\n是否开始？",
                "建立语义索引",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);
            if (confirm != DialogResult.OK)
            {
                return;
            }

            // The model is required for every mail. Check it up front so a
            // fresh machine fails fast with an actionable message instead of
            // reporting thousands of individual failures.
            SemanticServiceStatus model = addIn.SemanticClient.GetModelStatus();
            if (!model.Reachable)
            {
                MessageBox.Show(
                    this,
                    string.IsNullOrWhiteSpace(model.Message)
                        ? "本地语义搜索服务未运行。请先启动 embedding_service。"
                        : model.Message,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!model.ModelAvailable)
            {
                string detail = string.IsNullOrWhiteSpace(model.ModelError)
                    ? "本地语义模型尚未就绪。"
                    : model.ModelError;
                MessageBox.Show(
                    this,
                    detail + "\r\n\r\n请前往“设置”页面点击“下载语义搜索模型”（需联网，之后可离线使用），"
                        + "完成后再建立索引。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SetSemanticBusy(true, "正在建立语义索引…");
            _indexCancellation = new CancellationTokenSource();
            _semanticStopButton.Enabled = true;
            try
            {
                SemanticIndexService service = addIn.SemanticIndexService;
                SemanticIndexReport report =
                    await service.BuildIndexAsync(
                        addIn.Settings,
                        _indexCancellation.Token,
                        delegate(SemanticIndexProgress progress)
                        {
                            UpdateIndexProgress(progress);
                        },
                        delegate(Func<object> action)
                        {
                            return Invoke(action);
                        });
                ReportIndexOutcome(report);
            }
            catch (Exception ex)
            {
                ShowError("建立语义索引失败", ex);
            }
            finally
            {
                _semanticStopButton.Enabled = false;
                _indexCancellation.Dispose();
                _indexCancellation = null;
                SetSemanticBusy(false, null);
            }
        }

        private void ReportIndexOutcome(SemanticIndexReport report)
        {
            if (report.ModelUnavailable)
            {
                _semanticStatus.ForeColor = Color.FromArgb(176, 0, 32);
                _semanticStatus.Text = "索引已停止：本地语义模型不可用。\r\n"
                    + report.Error
                    + "\r\n请前往“设置”页面点击“下载语义搜索模型”，完成后再建立索引。";
            }
            else if (report.Cancelled)
            {
                _semanticStatus.ForeColor = Color.DimGray;
                _semanticStatus.Text = "索引已停止。已完成的 "
                    + report.Indexed + " 封邮件已保存，可稍后继续。";
            }
            else if (!string.IsNullOrWhiteSpace(report.Error))
            {
                _semanticStatus.ForeColor = Color.FromArgb(176, 0, 32);
                _semanticStatus.Text = "索引中止：" + report.Error;
            }
            else
            {
                _semanticStatus.ForeColor = Color.DimGray;
                _semanticStatus.Text = "索引完成：新增/更新 "
                    + report.Indexed + "，跳过未变化 " + report.Skipped
                    + "，失败 " + report.Failed + "。";
            }
        }

        private void SemanticStopClicked(object sender, EventArgs e)
        {
            if (_indexCancellation == null || _indexCancellation.IsCancellationRequested)
            {
                return;
            }

            _semanticStopButton.Enabled = false;
            SetStatus("正在停止索引…（当前这一封完成后停止）");
            _indexCancellation.Cancel();
        }

        private async void SemanticDownloadModelClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            SemanticServiceStatus status = addIn.SemanticClient.GetStatus();
            if (!status.Reachable)
            {
                MessageBox.Show(
                    this,
                    "本地语义搜索服务未运行。请先运行 embedding_service "
                    + "中的 start-service.cmd。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SetSemanticBusy(true, "正在下载本地语义模型，首次约 1.5 GB，请稍候…");
            try
            {
                SemanticServiceStatus result = await Task.Factory.StartNew(
                    delegate { return DownloadModel(addIn); },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                if (result.ModelAvailable)
                {
                    SetStatus("本地语义模型已就绪。");
                }
                else
                {
                    ShowError("下载模型失败", new InvalidOperationException(
                        string.IsNullOrWhiteSpace(result.ModelError)
                            ? "未知错误。"
                            : result.ModelError));
                }
            }
            finally
            {
                SetSemanticBusy(false, null);
                RefreshSemanticIndexStatus();
            }
        }

        private static SemanticServiceStatus DownloadModel(Connect addIn)
        {
            try
            {
                addIn.SemanticClient.DownloadModel();
            }
            catch (SemanticServiceException)
            {
                // Fall through: the status call below reports the reason.
            }

            return addIn.SemanticClient.GetModelStatus();
        }

        private void UpdateIndexProgress(SemanticIndexProgress progress)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<SemanticIndexProgress>(
                    UpdateIndexProgress), progress);
                return;
            }

            _semanticIndexStatus.Text = "索引中："
                + progress.Processed + " / " + progress.Total
                + "（新增 " + progress.Indexed + "，跳过 " + progress.Skipped
                + "，失败 " + progress.Failed + "）"
                + (string.IsNullOrWhiteSpace(progress.CurrentSubject)
                    ? string.Empty
                    : "\r\n" + DisplaySubject(progress.CurrentSubject));
            SetStatus("正在建立语义索引…Outlook 仍可继续使用。");
        }

        private void CheckServiceClicked(object sender, EventArgs e)
        {
            RefreshSemanticIndexStatus();
        }

        private async void SemanticStartServiceClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            if (addIn.SemanticServiceLauncher.IsRunning())
            {
                RefreshSemanticIndexStatus();
                SetStatus("本地语义搜索服务已在运行。");
                return;
            }

            if (!ShowInstallHintIfMissing())
            {
                return;
            }

            string startError = addIn.SemanticServiceLauncher.Start();
            if (!string.IsNullOrWhiteSpace(startError))
            {
                ShowServiceError(startError);
                return;
            }

            await WaitForServiceAsync(addIn, "本地语义搜索服务已启动。");
        }

        private async void SemanticRestartServiceClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            if (!ShowInstallHintIfMissing())
            {
                return;
            }

            SetSemanticBusy(true, "正在重启本地语义搜索服务…");
            try
            {
                string error = await Task.Factory.StartNew(
                    delegate { return addIn.SemanticServiceLauncher.Restart(); },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                if (!string.IsNullOrWhiteSpace(error))
                {
                    ShowServiceError(error);
                    return;
                }

                SetStatus("本地语义搜索服务已重启。");
            }
            finally
            {
                SetSemanticBusy(false, null);
                RefreshSemanticIndexStatus();
            }
        }

        private async void SemanticStopServiceClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            SetSemanticBusy(true, "正在停止本地语义搜索服务…");
            try
            {
                string error = await Task.Factory.StartNew(
                    delegate { return addIn.SemanticServiceLauncher.Stop(); },
                    CancellationToken.None,
                    TaskCreationOptions.LongRunning,
                    TaskScheduler.Default);
                if (!string.IsNullOrWhiteSpace(error))
                {
                    ShowServiceError(error);
                    return;
                }

                if (addIn.SemanticTray != null)
                {
                    addIn.SemanticTray.UpdateTrayIcon(false);
                }
                SetStatus("本地语义搜索服务已停止。");
            }
            finally
            {
                SetSemanticBusy(false, null);
                RefreshSemanticIndexStatus();
            }
        }

        private bool ShowInstallHintIfMissing()
        {
            if (SemanticServiceLauncher.IsInstalled())
            {
                return true;
            }

            MessageBox.Show(
                this,
                "尚未安装本地语义搜索服务。\r\n\r\n"
                + "请先解压发布包，在 embedding_service 目录中运行 "
                + "install-service.cmd 完成安装，然后回到这里点击“启动服务”。",
                "AI 邮件助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }

        private void ShowServiceError(string message)
        {
            SetStatus(message);
            MessageBox.Show(
                this,
                message,
                "AI 邮件助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Waits for the service to answer. Loading PyTorch takes a while, so
        /// poll for up to about a minute before giving up.
        /// </summary>
        private async Task WaitForServiceAsync(Connect addIn, string successMessage)
        {
            // Loading a 1.4 GB model on CPU can take a couple of minutes, so
            // allow up to three minutes before giving up. Elapsed time is
            // shown so the wait does not look like a hang.
            const int maximumSeconds = 180;
            SetSemanticBusy(true, null);
            try
            {
                for (int elapsed = 1; elapsed <= maximumSeconds; elapsed++)
                {
                    _semanticServiceStatus.Text = string.Format(
                        "服务状态：正在启动…（已等待 {0} 秒）", elapsed);
                    if (elapsed == 1)
                    {
                        SetStatus("正在启动本地语义搜索服务…首次加载模型可能需要 1-2 分钟。");
                    }

                    await Task.Delay(1000);
                    if (addIn.SemanticClient.GetStatus().Reachable)
                    {
                        if (addIn.SemanticTray != null)
                        {
                            addIn.SemanticTray.UpdateTrayIcon(true);
                        }
                        SetStatus(successMessage);
                        return;
                    }

                    // Only declare failure after at least 30 seconds if the
                    // background launcher process and python have definitely died.
                    // Checking too early (e.g. 5 seconds) causes false failures
                    // while PyTorch/CUDA are still importing in the background.
                    if (elapsed >= 30
                        && !addIn.SemanticServiceLauncher.IsServiceProcessAlive())
                    {
                        SetStatus("服务进程已退出，启动失败。请打开日志目录查看 "
                            + "semantic-service.log。设置页可一键打开日志目录。");
                        return;
                    }
                }

                SetStatus("服务在 3 分钟内未能响应。请打开日志目录查看 "
                    + "semantic-service.log 的最后几行。");
            }
            finally
            {
                SetSemanticBusy(false, null);
                RefreshSemanticIndexStatus();
            }
        }

        private void RefreshSemanticIndexStatus()
        {
            Connect addIn = Connect.Current;
            if (addIn == null)
            {
                _semanticServiceStatus.Text = "服务状态：加载项正在初始化…";
                _semanticIndexStatus.Text = string.Empty;
                return;
            }

            SemanticServiceStatus status = addIn.SemanticClient.GetStatus();
            if (addIn.SemanticTray != null)
            {
                addIn.SemanticTray.UpdateTrayIcon(status.Reachable);
            }

            if (!status.Reachable)
            {
                bool installed = SemanticServiceLauncher.IsInstalled();
                _semanticServiceStatus.Text = installed
                    ? "服务状态：未运行"
                    : "服务状态：未安装";
                _semanticServiceStatus.ForeColor = Color.FromArgb(176, 0, 32);
                _semanticIndexStatus.Text = installed
                    ? "点击“启动服务”即可开始，然后建立索引。"
                    : "请先在 embedding_service 目录运行 install-service.cmd。";
                _semanticIndexStatus.ForeColor = Color.FromArgb(176, 0, 32);
                return;
            }

            if (!status.ModelAvailable)
            {
                _semanticServiceStatus.Text = "服务状态：运行中；模型未就绪"
                    + FormatDeviceSuffix(status);
                _semanticServiceStatus.ForeColor = Color.FromArgb(176, 96, 0);
                _semanticIndexStatus.Text =
                    "请前往“设置”页面点击“下载语义搜索模型”（约 1.5 GB，需联网，之后可离线使用）。"
                    + (string.IsNullOrWhiteSpace(status.ModelError)
                        ? string.Empty
                        : "\r\n原因：" + status.ModelError);
                _semanticIndexStatus.ForeColor = Color.FromArgb(176, 96, 0);
                return;
            }

            _semanticServiceStatus.Text = "服务状态：运行中；模型就绪"
                + FormatDeviceSuffix(status);
            _semanticServiceStatus.ForeColor = Color.FromArgb(0, 120, 60);
            _semanticIndexStatus.Text =
                "索引：已索引 " + status.MailCount + " 封邮件，"
                + status.ChunkCount + " 个片段（"
                + status.EmbeddedChunkCount + " 个已向量化），占用 "
                + FormatSize(status.SizeBytes)
                + (string.IsNullOrWhiteSpace(status.LastSyncAt)
                    ? string.Empty
                    : "\r\n上次同步：" + status.LastSyncAt);
            _semanticIndexStatus.ForeColor = Color.DimGray;
        }

        /// <summary>
        /// Shows the compute device, which is auto-detected: an NVIDIA GPU
        /// when available, otherwise CPU.
        /// </summary>
        private static string FormatDeviceSuffix(SemanticServiceStatus status)
        {
            string device = string.IsNullOrWhiteSpace(status.Device)
                ? "未知"
                : status.Device;
            string reason = string.IsNullOrWhiteSpace(status.DeviceReason)
                ? string.Empty
                : "，" + status.DeviceReason;
            return "（" + device + reason + "）";
        }

        private void SetSemanticBusy(bool busy, string status)
        {
            _semanticBusy = busy;
            _semanticSearchButton.Enabled = !busy;
            _semanticIndexButton.Enabled = !busy;
            if (_semanticDownloadModelButton != null)
            {
                _semanticDownloadModelButton.Enabled = !busy;
            }
            _semanticStartServiceButton.Enabled = !busy;
            _semanticRestartServiceButton.Enabled = !busy;
            _semanticStopServiceButton.Enabled = !busy;
            _openServiceButton.Enabled = !busy;
            if (!busy)
            {
                _semanticStopButton.Enabled = false;
            }

            if (!string.IsNullOrWhiteSpace(status))
            {
                SetStatus(status);
            }
        }

        private void SelectSearchTab()
        {
            if (_tabs != null && _tabs.TabPages.Count > 1)
            {
                _tabs.SelectedIndex = 1;
            }
        }

        private void SelectSettingsTab()
        {
            if (_tabs != null && _tabs.TabPages.Count > 2)
            {
                _tabs.SelectedIndex = 2;
            }
        }

        private static string FormatSender(SemanticSearchResult result)
        {
            string name = (result.SenderName ?? string.Empty).Trim();
            string email = (result.SenderEmail ?? string.Empty).Trim();
            if (name.Length > 0 && email.Length > 0)
            {
                return name + " <" + email + ">";
            }

            return name.Length > 0 ? name : email;
        }

        private static string FormatDate(string value)
        {
            DateTime parsed;
            if (!DateTime.TryParse(
                value,
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AdjustToUniversal,
                out parsed))
            {
                return value ?? string.Empty;
            }

            return parsed.ToLocalTime().ToString("yyyy-MM-dd HH:mm");
        }

        private static string FormatSize(long bytes)
        {
            if (bytes < 1024)
            {
                return bytes + " B";
            }

            if (bytes < 1024 * 1024)
            {
                return (bytes / 1024.0).ToString("0.0") + " KB";
            }

            return (bytes / (1024.0 * 1024.0)).ToString("0.0") + " MB";
        }

        private void LoadSettingsToUI()
        {
            Connect addIn = Connect.Current;
            if (addIn == null || addIn.Settings == null)
            {
                return;
            }

            AppSettings settings = addIn.Settings;
            if (_settingsBaseUrl != null)
            {
                _settingsBaseUrl.Text = settings.ApiBaseUrl ?? string.Empty;
            }
            if (_settingsModel != null)
            {
                _settingsModel.Text = settings.Model ?? string.Empty;
            }
            if (_settingsReasoningEffort != null)
            {
                _settingsReasoningEffort.SelectedIndex = string.Equals(
                    settings.ReasoningEffort,
                    "medium",
                    StringComparison.OrdinalIgnoreCase)
                    ? 1
                    : (string.Equals(
                        settings.ReasoningEffort,
                        "max",
                        StringComparison.OrdinalIgnoreCase) ? 2 : 0);
            }
            if (_settingsApiKey != null)
            {
                _settingsApiKey.Text = string.Empty;
            }
            if (_settingsClearApiKey != null)
            {
                _settingsClearApiKey.Checked = false;
            }
            if (_settingsLanguage != null)
            {
                _settingsLanguage.Text = string.IsNullOrWhiteSpace(settings.SummaryLanguage)
                    ? "简体中文"
                    : settings.SummaryLanguage;
            }
            if (_settingsMaxCharacters != null)
            {
                _settingsMaxCharacters.Value = Math.Max(
                    _settingsMaxCharacters.Minimum,
                    Math.Min(
                        _settingsMaxCharacters.Maximum,
                        settings.MaxEmailCharacters > 0 ? settings.MaxEmailCharacters : 30000));
            }
            if (_settingsIdentityEmails != null)
            {
                _settingsIdentityEmails.Text = settings.IdentityEmailAddresses ?? string.Empty;
            }
            if (_settingsIdentityAliases != null)
            {
                _settingsIdentityAliases.Text = settings.IdentityAliases ?? string.Empty;
            }

            RefreshSettingsStatus();
        }

        private void SaveSettingsClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || addIn.Settings == null)
            {
                return;
            }

            try
            {
                AppSettings updated = addIn.Settings.Copy();
                updated.ProviderId = string.Empty;
                updated.ApiBaseUrl = (_settingsBaseUrl.Text ?? string.Empty).Trim();
                updated.Model = (_settingsModel.Text ?? string.Empty).Trim();
                updated.ReasoningEffort = (_settingsReasoningEffort.Text ?? "None").ToLowerInvariant();

                OpenAiCompatibleClient.BuildChatCompletionsUrl(updated.ApiBaseUrl);
                if (string.IsNullOrWhiteSpace(updated.Model))
                {
                    MessageBox.Show(
                        this,
                        "请填写 Model。",
                        "AI 邮件助手",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                bool endpointChanged = !SettingsStore.IsSameApiEndpoint(
                    addIn.Settings.ApiBaseUrl,
                    updated.ApiBaseUrl);
                if ((!addIn.SettingsStore.HasApiKey(addIn.Settings) || endpointChanged)
                    && string.IsNullOrWhiteSpace(_settingsApiKey.Text)
                    && !_settingsClearApiKey.Checked)
                {
                    MessageBox.Show(
                        this,
                        endpointChanged
                            ? "更改服务地址后，请重新输入 API Key。"
                            : "请填写 API Key。",
                        "AI 邮件助手",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                updated.SummaryLanguage = string.IsNullOrWhiteSpace(_settingsLanguage.Text)
                    ? "简体中文"
                    : _settingsLanguage.Text.Trim();
                updated.MaxEmailCharacters = Decimal.ToInt32(_settingsMaxCharacters.Value);
                updated.IdentityEmailAddresses = (_settingsIdentityEmails.Text ?? string.Empty).Trim();
                updated.IdentityAliases = (_settingsIdentityAliases.Text ?? string.Empty).Trim();

                if (_settingsClearApiKey.Checked)
                {
                    addIn.SettingsStore.SetApiKey(updated, string.Empty);
                }
                else if (!string.IsNullOrWhiteSpace(_settingsApiKey.Text))
                {
                    addIn.SettingsStore.SetApiKey(updated, _settingsApiKey.Text);
                }

                addIn.SettingsStore.Save(updated);
                addIn.ApplySettings(updated);

                _settingsApiKey.Text = string.Empty;
                _settingsClearApiKey.Checked = false;

                RefreshSettingsStatus();
                _settingsSaveStatus.Text = "设置已保存 (" + DateTime.Now.ToString("HH:mm:ss") + ")";
                _settingsSaveStatus.ForeColor = Color.FromArgb(0, 120, 60);
                SetStatus("设置已保存。");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "保存设置失败：" + ex.Message,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OpenDirectorySafely(string path)
        {
            string error = SemanticServiceLauncher.OpenDirectory(path);
            if (!string.IsNullOrWhiteSpace(error))
            {
                MessageBox.Show(
                    this,
                    error,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private bool EnsureConfigured()
        {
            Connect addIn = Connect.Current;
            if (addIn == null)
            {
                MessageBox.Show(
                    this,
                    "加载项尚未完成初始化，请重启 Outlook 后再试。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return false;
            }

            AppSettings settings = addIn.Settings;
            bool configured = settings != null
                && addIn.SettingsStore.HasApiKey(settings)
                && !string.IsNullOrWhiteSpace(settings.ApiBaseUrl)
                && !string.IsNullOrWhiteSpace(settings.Model);
            if (configured)
            {
                return true;
            }

            MessageBox.Show(
                this,
                "首次使用前，请前往“设置”页面填写 Base URL、API Key 和 Model 并保存。",
                "AI 邮件助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            SelectSettingsTab();
            return false;
        }

        private void RefreshSettingsStatus()
        {
            Connect addIn = Connect.Current;
            if (addIn == null || addIn.Settings == null)
            {
                if (_configurationStatus != null)
                {
                    _configurationStatus.Text = "加载项正在初始化……";
                }
                return;
            }

            AppSettings settings = addIn.Settings;
            if (_configurationStatus != null)
            {
                bool hasKey = addIn.SettingsStore.HasApiKey(settings);
                bool hasModel = !string.IsNullOrWhiteSpace(settings.Model);
                _configurationStatus.Text =
                    "当前状态：模型 " + (hasModel ? settings.Model : "未填写")
                    + " | API Key " + (hasKey ? "已加密保存" : "未配置");
                _configurationStatus.ForeColor = (hasKey && hasModel)
                    ? Color.FromArgb(0, 120, 60)
                    : Color.FromArgb(176, 0, 32);
            }
        }

        private void SetBusy(bool busy, string status)
        {
            _busy = busy;
            _summarizeButton.Enabled = !busy;
            if (!string.IsNullOrWhiteSpace(status))
            {
                SetStatus(status);
            }
        }

        private void SetStatus(string value)
        {
            _status.Text = value;
        }

        private void ShowError(string title, Exception exception)
        {
            Logger.Error(title, exception);
            SetStatus(title + "。");
            MessageBox.Show(
                this,
                title + "：" + exception.Message,
                "AI 邮件助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }

        private static Connect RequireAddIn()
        {
            if (Connect.Current == null)
            {
                throw new InvalidOperationException("加载项尚未完成初始化。");
            }

            return Connect.Current;
        }

        private static Button CreatePrimaryButton(string text)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = true;
            button.MinimumSize = new Size(180, 34);
            button.BackColor = Color.FromArgb(0, 103, 184);
            button.ForeColor = Color.White;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            return button;
        }



        private static string DisplaySubject(string subject)
        {
            string value = string.IsNullOrWhiteSpace(subject)
                ? "（无主题）"
                : subject.Trim();
            return value.Length <= 80 ? value : value.Substring(0, 80) + "…";
        }
    }
}
