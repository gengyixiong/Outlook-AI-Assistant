using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Diagnostics;
using OutlookAiAssistant.Models;
using OutlookAiAssistant.OutlookIntegration;

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
        private Button _summarizeButton;
        private Label _mailInfo;
        private RichTextBox _summaryOutput;
        private Button _copySummaryButton;
        private TextBox _searchInput;
        private Button _searchButton;
        private RichTextBox _searchPlanOutput;
        private Label _status;
        private Label _providerStatus;
        private bool _busy;

        public AssistantPaneControl()
        {
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            BuildInterface();
            Load += delegate { RefreshSettingsStatus(); };
        }

        private void BuildInterface()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 4;
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(root);

            Panel header = new Panel();
            header.Dock = DockStyle.Top;
            header.Height = 70;
            header.Padding = new Padding(16, 14, 16, 6);
            header.BackColor = Color.FromArgb(243, 247, 252);
            Label title = new Label();
            title.Text = "AI 邮件助手";
            title.Font = new Font("Segoe UI Semibold", 14F);
            title.AutoSize = true;
            title.Location = new Point(14, 12);
            Label subtitle = new Label();
            subtitle.Text = "手动摘要 · 本地 Outlook 搜索";
            subtitle.ForeColor = Color.DimGray;
            subtitle.AutoSize = true;
            subtitle.Location = new Point(16, 43);
            header.Controls.Add(title);
            header.Controls.Add(subtitle);
            root.Controls.Add(header);

            Label privacy = new Label();
            privacy.Text =
                "摘要：仅点击按钮后读取并发送当前邮件及同一会话历史。\r\n"
                + "搜索：只发送你输入的描述，邮件与结果始终留在本机。";
            privacy.AutoSize = true;
            privacy.MaximumSize = new Size(380, 0);
            privacy.Margin = new Padding(16, 10, 16, 8);
            privacy.ForeColor = Color.FromArgb(45, 82, 120);
            root.Controls.Add(privacy);

            TabControl tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            tabs.Margin = new Padding(10, 0, 10, 4);
            tabs.TabPages.Add(BuildSummaryPage());
            tabs.TabPages.Add(BuildSearchPage());
            tabs.TabPages.Add(BuildSettingsPage());
            root.Controls.Add(tabs);

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

            _summarizeButton = CreatePrimaryButton("生成会话背景与当前邮件摘要");
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

        private TabPage BuildSearchPage()
        {
            TabPage page = new TabPage("智能搜索");
            page.Padding = new Padding(10);
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 4;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            Label prompt = new Label();
            prompt.Text =
                "描述你要找的邮件，例如：\r\n"
                + "找 Richard 上个月发给我的、关于 Malaysia 报价且带附件的邮件";
            prompt.AutoSize = true;
            prompt.MaximumSize = new Size(350, 0);
            layout.Controls.Add(prompt);

            _searchInput = new TextBox();
            _searchInput.Multiline = true;
            _searchInput.ScrollBars = ScrollBars.Vertical;
            _searchInput.Height = 78;
            _searchInput.Dock = DockStyle.Top;
            _searchInput.Margin = new Padding(3, 8, 3, 8);
            layout.Controls.Add(_searchInput);

            _searchButton = CreatePrimaryButton("解析并在 Outlook 中搜索");
            _searchButton.Click += SearchClicked;
            layout.Controls.Add(_searchButton);

            _searchPlanOutput = new RichTextBox();
            _searchPlanOutput.Dock = DockStyle.Fill;
            _searchPlanOutput.ReadOnly = true;
            _searchPlanOutput.BackColor = Color.White;
            _searchPlanOutput.BorderStyle = BorderStyle.FixedSingle;
            _searchPlanOutput.Margin = new Padding(3, 10, 3, 3);
            layout.Controls.Add(_searchPlanOutput);
            page.Controls.Add(layout);
            return page;
        }

        private TabPage BuildSettingsPage()
        {
            TabPage page = new TabPage("设置");
            page.Padding = new Padding(12);
            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 3;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            _providerStatus = new Label();
            _providerStatus.AutoSize = true;
            _providerStatus.MaximumSize = new Size(350, 0);
            layout.Controls.Add(_providerStatus);

            Button settingsButton = CreatePrimaryButton("打开 API 设置");
            settingsButton.Margin = new Padding(3, 12, 3, 12);
            settingsButton.Click += delegate
            {
                Connect addIn = Connect.Current;
                if (addIn != null && addIn.ShowSettings(this))
                {
                    RefreshSettingsStatus();
                    SetStatus("设置已保存。");
                }
            };
            layout.Controls.Add(settingsButton);

            Label note = new Label();
            note.AutoSize = true;
            note.MaximumSize = new Size(350, 0);
            note.ForeColor = Color.DimGray;
            note.Text =
                "API Key 加密保存在当前 Windows 用户目录中；不会写入项目、"
                + "安装目录或日志。更换电脑后需要重新填写密钥。";
            layout.Controls.Add(note);
            page.Controls.Add(layout);
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
                SetStatus("正在发送会话历史和当前邮件并生成摘要……");
                string summary = await addIn.SummaryService.SummarizeAsync(
                    conversation,
                    addIn.Settings,
                    CancellationToken.None);
                _summaryOutput.Text = summary;
                _copySummaryButton.Enabled = true;
                SetStatus("会话背景与当前邮件摘要已生成。");
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

        private async void SearchClicked(object sender, EventArgs e)
        {
            if (_busy || !EnsureConfigured())
            {
                return;
            }

            string description = _searchInput.Text.Trim();
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

            SetBusy(true, "正在把描述转换成搜索条件……");
            try
            {
                Connect addIn = RequireAddIn();
                // Only the description string is sent to the configured AI service.
                SearchPlan plan = await addIn.SearchPlanner.CreatePlanAsync(
                    description,
                    addIn.Settings,
                    CancellationToken.None);
                string aqs = addIn.QueryCompiler.Compile(plan);
                _searchPlanOutput.Text =
                    addIn.SearchPlanFormatter.Format(plan, aqs);
                SetStatus("正在调用 Outlook 本地搜索……");
                addIn.OutlookSearch.Search(aqs, plan.Scope);
                SetStatus("搜索条件已交给 Outlook，本地结果已显示。");
            }
            catch (Exception ex)
            {
                ShowError("搜索失败", ex);
            }
            finally
            {
                SetBusy(false, null);
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
                "首次使用前，请先配置 API 地址、模型和 API Key。",
                "AI 邮件助手",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            bool saved = addIn.ShowSettings(this);
            RefreshSettingsStatus();
            return saved;
        }

        private void RefreshSettingsStatus()
        {
            Connect addIn = Connect.Current;
            if (addIn == null || addIn.Settings == null)
            {
                _providerStatus.Text = "加载项正在初始化……";
                return;
            }

            AppSettings settings = addIn.Settings;
            _providerStatus.Text =
                "提供商：" + settings.ProviderId + "\r\n"
                + "模型：" + (string.IsNullOrWhiteSpace(settings.Model)
                    ? "未配置"
                    : settings.Model) + "\r\n"
                + "API Key：" + (addIn.SettingsStore.HasApiKey(settings)
                    ? "已加密保存"
                    : "未配置");
        }

        private void SetBusy(bool busy, string status)
        {
            _busy = busy;
            _summarizeButton.Enabled = !busy;
            _searchButton.Enabled = !busy;
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
