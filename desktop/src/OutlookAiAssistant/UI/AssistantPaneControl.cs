using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
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
        private Button _openServiceButton;
        private Label _semanticServiceStatus;
        private Label _semanticIndexStatus;
        private FlowLayoutPanel _semanticResultsPanel;
        private Label _semanticStatus;
        private bool _semanticBusy;
        private CancellationTokenSource _indexCancellation;

        public AssistantPaneControl()
        {
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            BuildInterface();
            Load += delegate
            {
                RefreshSettingsStatus();
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

            TabControl tabs = new TabControl();
            tabs.Dock = DockStyle.Fill;
            tabs.Margin = new Padding(10, 8, 10, 4);
            tabs.TabPages.Add(BuildSummaryPage());
            tabs.TabPages.Add(BuildSemanticSearchPage());
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
            layout.RowCount = 8;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
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
            _semanticServiceStatus.MaximumSize = new Size(350, 0);
            _semanticServiceStatus.Font = new Font(Font, FontStyle.Bold);
            _semanticServiceStatus.Text = "服务状态：正在检查…";
            layout.Controls.Add(_semanticServiceStatus);

            Label prompt = new Label();
            prompt.Text =
                "用中文描述你要找的邮件内容，例如：\r\n"
                + "之前印度客户问过幕墙玻璃 Panel 能不能单独编辑的邮件";
            prompt.AutoSize = true;
            prompt.MaximumSize = new Size(350, 0);
            layout.Controls.Add(prompt);

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

            _semanticDownloadModelButton = new Button();
            _semanticDownloadModelButton.Text = "下载模型";
            _semanticDownloadModelButton.AutoSize = true;
            _semanticDownloadModelButton.MinimumSize = new Size(88, 29);
            _semanticDownloadModelButton.Margin = new Padding(3, 0, 3, 3);
            _semanticDownloadModelButton.Click += SemanticDownloadModelClicked;
            actions.Controls.Add(_semanticDownloadModelButton);

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
            _semanticIndexStatus.MaximumSize = new Size(350, 0);
            _semanticIndexStatus.ForeColor = Color.DimGray;
            _semanticIndexStatus.Text = "索引状态：未检查";
            layout.Controls.Add(_semanticIndexStatus);

            _semanticStatus = new Label();
            _semanticStatus.AutoSize = true;
            _semanticStatus.MaximumSize = new Size(350, 0);
            _semanticStatus.Text = string.Empty;
            layout.Controls.Add(_semanticStatus);

            _semanticResultsPanel = new FlowLayoutPanel();
            _semanticResultsPanel.Dock = DockStyle.Fill;
            _semanticResultsPanel.FlowDirection = FlowDirection.TopDown;
            _semanticResultsPanel.WrapContents = false;
            _semanticResultsPanel.AutoScroll = true;
            _semanticResultsPanel.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_semanticResultsPanel);

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

            _configurationStatus = new Label();
            _configurationStatus.AutoSize = true;
            _configurationStatus.MaximumSize = new Size(350, 0);
            layout.Controls.Add(_configurationStatus);

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
                    _semanticResultsPanel.Controls.Add(
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

        private void AddSemanticFallbackControls(SemanticSearchFacade facade)
        {
            Button checkService = new Button();
            checkService.Text = "检查本地服务状态";
            checkService.AutoSize = true;
            checkService.Click += CheckServiceClicked;
            _semanticResultsPanel.Controls.Add(checkService);

            Button outlookSearch = new Button();
            outlookSearch.Text = "改用 Outlook 原生搜索";
            outlookSearch.AutoSize = true;
            outlookSearch.Click += delegate
            {
                SelectSearchTab();
            };
            _semanticResultsPanel.Controls.Add(outlookSearch);
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
        /// the layout can be unit tested; a previous version used nested
        /// Dock=Top panels that collapsed to an 18-pixel-wide sliver inside
        /// the scrolling results panel, which made every result invisible.
        /// </summary>
        public static Control BuildResultCardControl(
            SemanticSearchResult result,
            int rank,
            Font baseFont,
            Action openAction,
            Action outlookAction)
        {
            // A single-column TableLayoutPanel with AutoSize rows is the most
            // reliable layout inside an auto-scrolling FlowLayoutPanel.
            const int cardWidth = 330;
            const int textWidth = cardWidth - 28;

            TableLayoutPanel card = new TableLayoutPanel();
            card.ColumnCount = 1;
            card.AutoSize = true;
            card.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            // Fix the width so labels wrap predictably; the panel still grows
            // vertically with the content.
            card.Width = cardWidth;
            card.MinimumSize = new Size(cardWidth, 0);
            card.MaximumSize = new Size(cardWidth, 0);
            card.Margin = new Padding(3, 4, 3, 4);
            card.Padding = new Padding(8);
            card.BackColor = Color.White;
            card.CellBorderStyle = TableLayoutPanelCellBorderStyle.None;

            Label title = new Label();
            title.Text = rank + ". " + DisplaySubjectStatic(result.Subject);
            title.Font = new Font(baseFont, FontStyle.Bold);
            title.AutoSize = true;
            title.MaximumSize = new Size(textWidth, 0);
            title.Margin = new Padding(0, 0, 0, 2);
            card.Controls.Add(title);

            Label meta = new Label();
            meta.Text = FormatSender(result) + "\r\n"
                + FormatDate(result.ReceivedAt)
                + (string.IsNullOrWhiteSpace(result.FolderPath)
                    ? string.Empty
                    : "  ·  " + result.FolderPath);;
            meta.AutoSize = true;
            meta.MaximumSize = new Size(textWidth, 0);
            meta.ForeColor = Color.DimGray;
            meta.Margin = new Padding(0, 0, 0, 4);
            card.Controls.Add(meta);

            if (!string.IsNullOrWhiteSpace(result.AttachmentNames))
            {
                Label attachments = new Label();
                attachments.Text = "附件：" + result.AttachmentNames;
                attachments.AutoSize = true;
                attachments.MaximumSize = new Size(textWidth, 0);
                attachments.ForeColor = Color.DimGray;
                attachments.Margin = new Padding(0, 0, 0, 4);
                card.Controls.Add(attachments);
            }

            Label snippet = new Label();
            snippet.Text = string.IsNullOrWhiteSpace(result.Snippet)
                ? "（无正文片段）"
                : result.Snippet;
            snippet.AutoSize = true;
            snippet.MaximumSize = new Size(textWidth, 0);
            snippet.Margin = new Padding(0, 0, 0, 6);
            card.Controls.Add(snippet);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.AutoSize = true;
            buttons.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            buttons.WrapContents = false;
            buttons.Margin = new Padding(0);
            buttons.MinimumSize = new Size(0, 32);

            Button open = new Button();
            open.Text = "打开邮件";
            open.AutoSize = true;
            open.Click += delegate { openAction(); };
            buttons.Controls.Add(open);

            Button outlook = new Button();
            outlook.Text = "Outlook 搜索";
            outlook.AutoSize = true;
            outlook.Click += delegate { outlookAction(); };
            buttons.Controls.Add(outlook);
            card.Controls.Add(buttons);

            // Give the card a visible outline without a BorderStyle, which
            // does not render on an AutoSize panel.
            Panel frame = new Panel();
            frame.BorderStyle = BorderStyle.FixedSingle;
            frame.AutoSize = true;
            frame.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            frame.Margin = new Padding(3, 4, 3, 4);
            frame.Padding = new Padding(1);
            frame.Controls.Add(card);
            card.Margin = new Padding(0);
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
                    detail + "\r\n\r\n请先点击“下载模型”（需要联网，之后可离线使用），"
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
                    + "\r\n请先点击“下载模型”，完成后再建立索引。";
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

            if (!SemanticServiceLauncher.IsInstalled())
            {
                MessageBox.Show(
                    this,
                    "尚未安装本地语义搜索服务。\r\n\r\n"
                    + "请先解压发布包，在 embedding_service 目录中运行 "
                    + "install-service.cmd 完成安装，然后回到这里点击“启动服务”。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            string startError = addIn.SemanticServiceLauncher.Start();
            if (!string.IsNullOrWhiteSpace(startError))
            {
                MessageBox.Show(
                    this,
                    startError,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            // The service needs a few seconds to import PyTorch. Poll briefly
            // so the status label reflects reality instead of staying "stopped".
            _semanticServiceStatus.Text = "服务状态：正在启动…";
            SetStatus("正在启动本地语义搜索服务…");
            SetSemanticBusy(true, null);
            try
            {
                for (int attempt = 0; attempt < 30; attempt++)
                {
                    await Task.Delay(1000);
                    if (addIn.SemanticClient.GetStatus().Reachable)
                    {
                        SetStatus("本地语义搜索服务已启动。");
                        return;
                    }
                }

                SetStatus("服务启动较慢或失败，请查看弹出的服务窗口。");
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
                _semanticServiceStatus.Text = "服务状态：运行中；模型未就绪";
                _semanticServiceStatus.ForeColor = Color.FromArgb(176, 96, 0);
                _semanticIndexStatus.Text =
                    "点击“下载模型”完成一次下载（需要联网，之后可离线使用）。"
                    + (string.IsNullOrWhiteSpace(status.ModelError)
                        ? string.Empty
                        : "\r\n原因：" + status.ModelError);
                _semanticIndexStatus.ForeColor = Color.FromArgb(176, 96, 0);
                return;
            }

            _semanticServiceStatus.Text = "服务状态：运行中；模型就绪"
                + "（" + status.Device + "）";
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

        private void SetSemanticBusy(bool busy, string status)
        {
            _semanticBusy = busy;
            _semanticSearchButton.Enabled = !busy;
            _semanticIndexButton.Enabled = !busy;
            _semanticDownloadModelButton.Enabled = !busy;
            _semanticStartServiceButton.Enabled = !busy;
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
            Control parent = _semanticResultsPanel;
            while (parent != null && !(parent is TabControl))
            {
                parent = parent.Parent;
            }

            TabControl tabs = parent as TabControl;
            if (tabs != null && tabs.TabPages.Count > 1)
            {
                tabs.SelectedIndex = 1;
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
                "首次使用前，请先配置 Base URL、API Key 和 Model。",
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
                _configurationStatus.Text = "加载项正在初始化……";
                return;
            }

            AppSettings settings = addIn.Settings;
            _configurationStatus.Text =
                "模型：" + settings.Model + "\r\n"
                + "API Key：" + (addIn.SettingsStore.HasApiKey(settings)
                    ? "已加密保存"
                    : "未配置");
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
