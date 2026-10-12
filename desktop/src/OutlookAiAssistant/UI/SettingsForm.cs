using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.EntityIndex;
using OutlookAiAssistant.SemanticSearch;

namespace OutlookAiAssistant.UI
{
    /// <summary>
    /// User-editable OpenAI-compatible settings, summary identity and
    /// explicitly initiated Outlook contact-index rebuilds.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly SettingsStore _settingsStore;
        private AppSettings _settings;
        private readonly ContactIndexStore _contactIndexStore;
        private readonly ContactIndexService _contactIndexService;
        private TextBox _baseUrl;
        private TextBox _apiKey;
        private ComboBox _model;
        private ComboBox _reasoningEffort;
        private CheckBox _showApiKey;
        private CheckBox _clearApiKey;
        private ComboBox _language;
        private NumericUpDown _maximumCharacters;
        private TextBox _identityEmails;
        private TextBox _identityAliases;
        private Label _apiHelp;
        private Label _indexStatus;
        private Button _rebuildIndex;

        // Semantic index management.
        private Label _semanticStatus;
        private NumericUpDown _semanticMonths;
        private NumericUpDown _semanticMaxMails;
        private CheckBox _semanticIncludeSent;
        private Button _semanticSyncButton;
        private Button _semanticRebuildButton;
        private Button _semanticDeleteButton;
        private Button _semanticRefreshButton;
        private Button _semanticDownloadButton;
        private Button _semanticStopButton;
        private bool _semanticBusy;
        private CancellationTokenSource _semanticCancellation;

        public AppSettings SavedSettings { get; private set; }

        public SettingsForm(
            SettingsStore settingsStore,
            AppSettings settings)
            : this(settingsStore, settings, null, null)
        {
        }

        public SettingsForm(
            SettingsStore settingsStore,
            AppSettings settings,
            ContactIndexStore contactIndexStore,
            ContactIndexService contactIndexService)
        {
            _settingsStore = settingsStore;
            _settings = settings.Copy();
            _contactIndexStore = contactIndexStore;
            _contactIndexService = contactIndexService;
            Text = "AI 邮件助手设置";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(680, 600);
            Size = new Size(760, 730);
            Font = new Font("Segoe UI", 9F);
            ShowIcon = false;
            BuildInterface();
            LoadSettings();
            RefreshIndexStatus();
            RefreshSemanticStatus();
        }

        private void BuildInterface()
        {
            Panel scroll = new Panel();
            scroll.Dock = DockStyle.Fill;
            scroll.AutoScroll = true;
            Controls.Add(scroll);

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Top;
            root.AutoSize = true;
            root.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            root.Padding = new Padding(18);
            root.ColumnCount = 2;
            root.RowCount = 15;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            scroll.Controls.Add(root);

            _baseUrl = CreateTextBox();
            AddRow(root, 0, "Base URL", _baseUrl);

            TableLayoutPanel apiKeyPanel = new TableLayoutPanel();
            apiKeyPanel.Dock = DockStyle.Top;
            apiKeyPanel.AutoSize = true;
            apiKeyPanel.ColumnCount = 1;
            _apiKey = CreateTextBox();
            _apiKey.UseSystemPasswordChar = true;
            apiKeyPanel.Controls.Add(_apiKey);

            FlowLayoutPanel keyOptions = new FlowLayoutPanel();
            keyOptions.AutoSize = true;
            keyOptions.Dock = DockStyle.Top;
            _showApiKey = new CheckBox();
            _showApiKey.AutoSize = true;
            _showApiKey.Text = "显示本次输入";
            _showApiKey.CheckedChanged += delegate
            {
                _apiKey.UseSystemPasswordChar = !_showApiKey.Checked;
            };
            _clearApiKey = new CheckBox();
            _clearApiKey.AutoSize = true;
            _clearApiKey.Text = "清除已保存的密钥";
            keyOptions.Controls.Add(_showApiKey);
            keyOptions.Controls.Add(_clearApiKey);
            apiKeyPanel.Controls.Add(keyOptions);
            AddRow(root, 1, "API Key", apiKeyPanel);

            Label keyNote = CreateNoteLabel();
            keyNote.Text =
                "留空表示继续使用已保存的密钥。密钥使用 Windows DPAPI 加密，"
                + "不会在窗口中回显。更改服务地址后必须重新输入 API Key。";
            root.Controls.Add(keyNote, 1, 2);

            _model = CreateComboBox();
            _model.DropDownStyle = ComboBoxStyle.DropDown;
            _model.Items.Add("gpt-5.6-luna");
            AddRow(root, 3, "Model", _model);

            _reasoningEffort = CreateComboBox();
            _reasoningEffort.Items.AddRange(new object[] { "None", "Medium", "Max" });
            AddRow(root, 4, "Reasoning Effort", _reasoningEffort);

            _apiHelp = CreateNoteLabel();
            _apiHelp.Text = "支持 OpenAI-compatible Chat Completions API。模型名称可以手动输入。";
            root.Controls.Add(_apiHelp, 1, 5);

            _language = CreateComboBox();
            _language.Items.AddRange(
                new object[] { "简体中文", "English", "跟随邮件语言" });
            AddRow(root, 6, "摘要语言", _language);

            _maximumCharacters = new NumericUpDown();
            _maximumCharacters.Dock = DockStyle.Top;
            _maximumCharacters.Minimum = 5000;
            _maximumCharacters.Maximum = 200000;
            _maximumCharacters.Increment = 5000;
            _maximumCharacters.ThousandsSeparator = true;
            AddRow(root, 7, "正文字符上限", _maximumCharacters);

            _identityEmails = CreateMultilineTextBox(64);
            AddRow(root, 8, "我的邮箱地址", _identityEmails);

            _identityAliases = CreateMultilineTextBox(82);
            AddRow(root, 9, "别人对我的称呼", _identityAliases);

            _indexStatus = new Label();
            _indexStatus.AutoSize = true;
            _indexStatus.MaximumSize = new Size(520, 0);
            AddRow(root, 10, "联系人索引", _indexStatus);

            FlowLayoutPanel rebuildPanel = new FlowLayoutPanel();
            rebuildPanel.Dock = DockStyle.Top;
            rebuildPanel.AutoSize = true;
            rebuildPanel.FlowDirection = FlowDirection.TopDown;
            _rebuildIndex = new Button();
            _rebuildIndex.Text = "重建联系人索引";
            _rebuildIndex.AutoSize = true;
            _rebuildIndex.Enabled = _contactIndexService != null;
            _rebuildIndex.Click += RebuildIndexClicked;
            rebuildPanel.Controls.Add(_rebuildIndex);
            Label rebuildNote = CreateNoteLabel();
            rebuildNote.Text =
                "仅点击此按钮时扫描 Classic Outlook。Inbox 旁边及其下级的"
                + "自建邮件文件夹会作为人工归类标签，并关联其中邮件的联系人；"
                + "每人最早 1 封和最近最多 3 封代表邮件的必要片段会发送给"
                + "当前 AI；不会后台扫描。";
            rebuildPanel.Controls.Add(rebuildNote);
            AddRow(root, 11, string.Empty, rebuildPanel);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.AutoSize = true;
            Button save = new Button();
            save.Text = "保存";
            save.AutoSize = true;
            save.Click += SaveClicked;
            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.AutoSize = true;
            cancel.DialogResult = DialogResult.Cancel;
            buttons.Controls.Add(save);
            buttons.Controls.Add(cancel);

            // ---- Semantic index management ----
            _semanticStatus = new Label();
            _semanticStatus.AutoSize = true;
            _semanticStatus.MaximumSize = new Size(520, 0);
            AddRow(root, 13, "语义索引", BuildSemanticPanel());

            root.Controls.Add(buttons, 0, 12);
            root.SetColumnSpan(buttons, 2);

            FlowLayoutPanel bottomButtons = new FlowLayoutPanel();
            bottomButtons.Dock = DockStyle.Fill;
            bottomButtons.FlowDirection = FlowDirection.RightToLeft;
            bottomButtons.AutoSize = true;
            Button save2 = new Button();
            save2.Text = "保存";
            save2.AutoSize = true;
            save2.Click += SaveClicked;
            bottomButtons.Controls.Add(save2);
            root.Controls.Add(bottomButtons, 0, 14);
            root.SetColumnSpan(bottomButtons, 2);

            AcceptButton = save;
            CancelButton = cancel;
        }

        private Control BuildSemanticPanel()
        {
            TableLayoutPanel panel = new TableLayoutPanel();
            panel.ColumnCount = 1;
            panel.AutoSize = true;
            panel.Dock = DockStyle.Top;

            _semanticStatus.AutoSize = true;
            _semanticStatus.MaximumSize = new Size(520, 0);
            _semanticStatus.Text = "语义索引状态：未检查";
            panel.Controls.Add(_semanticStatus);

            TableLayoutPanel range = new TableLayoutPanel();
            range.ColumnCount = 4;
            range.AutoSize = true;
            range.Dock = DockStyle.Top;
            Label monthsLabel = new Label();
            monthsLabel.Text = "索引最近（月）";
            monthsLabel.AutoSize = true;
            monthsLabel.Anchor = AnchorStyles.Left;
            _semanticMonths = new NumericUpDown();
            _semanticMonths.Minimum = 1;
            _semanticMonths.Maximum = 120;
            _semanticMonths.Width = 70;
            Label maxLabel = new Label();
            maxLabel.Text = "最多邮件数";
            maxLabel.AutoSize = true;
            maxLabel.Anchor = AnchorStyles.Left;
            _semanticMaxMails = new NumericUpDown();
            _semanticMaxMails.Minimum = 100;
            _semanticMaxMails.Maximum = 100000;
            _semanticMaxMails.Increment = 500;
            _semanticMaxMails.ThousandsSeparator = true;
            _semanticMaxMails.Width = 90;
            range.Controls.Add(monthsLabel, 0, 0);
            range.Controls.Add(_semanticMonths, 1, 0);
            range.Controls.Add(maxLabel, 2, 0);
            range.Controls.Add(_semanticMaxMails, 3, 0);
            panel.Controls.Add(range);

            _semanticIncludeSent = new CheckBox();
            _semanticIncludeSent.AutoSize = true;
            _semanticIncludeSent.Text = "包含已发送邮件";
            panel.Controls.Add(_semanticIncludeSent);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.AutoSize = true;
            buttons.WrapContents = true;

            _semanticSyncButton = new Button();
            _semanticSyncButton.Text = "增量同步";
            _semanticSyncButton.AutoSize = true;
            _semanticSyncButton.Click += SemanticSyncClicked;
            buttons.Controls.Add(_semanticSyncButton);

            _semanticDownloadButton = new Button();
            _semanticDownloadButton.Text = "下载模型";
            _semanticDownloadButton.AutoSize = true;
            _semanticDownloadButton.Click += SemanticDownloadClicked;
            buttons.Controls.Add(_semanticDownloadButton);

            _semanticStopButton = new Button();
            _semanticStopButton.Text = "停止索引";
            _semanticStopButton.AutoSize = true;
            _semanticStopButton.Enabled = false;
            _semanticStopButton.Click += SemanticStopClicked;
            buttons.Controls.Add(_semanticStopButton);

            _semanticRebuildButton = new Button();
            _semanticRebuildButton.Text = "重新构建索引";
            _semanticRebuildButton.AutoSize = true;
            _semanticRebuildButton.Click += SemanticRebuildClicked;
            buttons.Controls.Add(_semanticRebuildButton);

            _semanticDeleteButton = new Button();
            _semanticDeleteButton.Text = "删除本地索引";
            _semanticDeleteButton.AutoSize = true;
            _semanticDeleteButton.Click += SemanticDeleteClicked;
            buttons.Controls.Add(_semanticDeleteButton);

            _semanticRefreshButton = new Button();
            _semanticRefreshButton.Text = "刷新状态";
            _semanticRefreshButton.AutoSize = true;
            _semanticRefreshButton.Click += delegate { RefreshSemanticStatus(); };
            buttons.Controls.Add(_semanticRefreshButton);
            panel.Controls.Add(buttons);

            Label note = CreateNoteLabel();
            note.Text =
                "语义索引读取邮件正文，并在本机生成向量，保存在 "
                + "%LOCALAPPDATA%\\OutlookAiAssistant\\semantic-index。"
                + "邮件正文和向量不会上传到任何云端服务。删除索引会同时删除"
                + "本机保存的邮件正文副本和向量。仍然使用 Outlook 原生搜索与"
                + "Executive Brief 不受影响。";
            panel.Controls.Add(note);
            return panel;
        }

        private void LoadSettings()
        {
            _baseUrl.Text = _settings.ApiBaseUrl;
            _model.Text = _settings.Model;
            _reasoningEffort.SelectedIndex = string.Equals(
                _settings.ReasoningEffort,
                "medium",
                StringComparison.OrdinalIgnoreCase)
                ? 1
                : (string.Equals(
                    _settings.ReasoningEffort,
                    "max",
                    StringComparison.OrdinalIgnoreCase) ? 2 : 0);
            _apiKey.Text = string.Empty;
            _clearApiKey.Checked = false;
            _language.Text = _settings.SummaryLanguage;
            _maximumCharacters.Value = Math.Max(
                _maximumCharacters.Minimum,
                Math.Min(
                    _maximumCharacters.Maximum,
                    _settings.MaxEmailCharacters));
            _identityEmails.Text = _settings.IdentityEmailAddresses;
            _identityAliases.Text = _settings.IdentityAliases;
            _semanticMonths.Value = Math.Max(
                _semanticMonths.Minimum,
                Math.Min(_semanticMonths.Maximum, _settings.SemanticIndexMonths));
            _semanticMaxMails.Value = Math.Max(
                _semanticMaxMails.Minimum,
                Math.Min(
                    _semanticMaxMails.Maximum, _settings.SemanticIndexMaxMails));
            _semanticIncludeSent.Checked = _settings.SemanticIndexIncludeSent;
        }

        private async void SemanticSyncClicked(object sender, EventArgs e)
        {
            await RunSemanticBuild(false);
        }

        private async void SemanticRebuildClicked(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                this,
                "重新构建会为本范围内所有邮件重新计算向量，可能耗时较长。"
                + "已保存的索引会在成功后才能被替换。是否继续？",
                "重新构建语义索引",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning) != DialogResult.OK)
            {
                return;
            }

            await RunSemanticBuild(true);
        }

        private async System.Threading.Tasks.Task RunSemanticBuild(
            bool forceReembed)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            AppSettings updated;
            if (!TrySaveSettings(out updated))
            {
                return;
            }

            updated.SemanticIndexMonths = (int)_semanticMonths.Value;
            updated.SemanticIndexMaxMails = (int)_semanticMaxMails.Value;
            updated.SemanticIndexIncludeSent = _semanticIncludeSent.Checked;
            updated = PersistWithoutClosing(updated);

            SemanticServiceStatus status = addIn.SemanticClient.GetStatus();
            if (!status.Reachable)
            {
                MessageBox.Show(
                    this,
                    "本地语义搜索服务未运行。请先运行 embedding_service "
                    + "中的 setup-venv.cmd 和 start-service.cmd。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            if (!status.ModelAvailable)
            {
                MessageBox.Show(
                    this,
                    "本地语义模型尚未就绪，无法建立索引。\r\n\r\n"
                    + "请先点击“下载模型”完成一次下载（需要联网，之后可离线使用）。"
                    + (string.IsNullOrWhiteSpace(status.ModelError)
                        ? string.Empty
                        : "\r\n\r\n原因：" + status.ModelError),
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SetSemanticBusy(true);
            _semanticStatus.Text = "正在枚举 Outlook 邮件…";
            _semanticCancellation = new CancellationTokenSource();
            _semanticStopButton.Enabled = true;
            try
            {
                SemanticIndexReport report =
                    await addIn.SemanticIndexService.BuildIndexAsync(
                        updated,
                        _semanticCancellation.Token,
                        delegate(SemanticIndexProgress progress)
                        {
                            UpdateSemanticProgress(progress);
                        },
                        delegate(Func<object> action)
                        {
                            return Invoke(action);
                        });
                if (report.ModelUnavailable)
                {
                    _semanticStatus.Text =
                        "索引已停止：本地语义模型不可用。请先点击“下载模型”。"
                        + "\r\n" + report.Error;
                }
                else if (report.Cancelled)
                {
                    _semanticStatus.Text = "索引已停止。已完成的 "
                        + report.Indexed + " 封邮件已保存，可稍后继续。";
                }
                else if (!string.IsNullOrWhiteSpace(report.Error))
                {
                    _semanticStatus.Text = "索引中止：" + report.Error;
                }
                else
                {
                    _semanticStatus.Text = "索引完成：新增/更新 "
                        + report.Indexed + "，跳过 " + report.Skipped
                        + "，失败 " + report.Failed + "。";
                }
            }
            catch (Exception ex)
            {
                _semanticStatus.Text = "索引失败：" + ex.Message;
            }
            finally
            {
                _semanticStopButton.Enabled = false;
                _semanticCancellation.Dispose();
                _semanticCancellation = null;
                SetSemanticBusy(false);
                RefreshSemanticStatus();
            }
        }

        private void SemanticStopClicked(object sender, EventArgs e)
        {
            if (_semanticCancellation == null
                || _semanticCancellation.IsCancellationRequested)
            {
                return;
            }

            _semanticStopButton.Enabled = false;
            _semanticStatus.Text = "正在停止索引…（当前这一封完成后停止）";
            _semanticCancellation.Cancel();
        }

        private async void SemanticDownloadClicked(object sender, EventArgs e)
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
                    + "中的 setup-venv.cmd 和 start-service.cmd。",
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            SetSemanticBusy(true);
            _semanticStatus.Text = "正在下载本地语义模型（首次约 1.5 GB）…";
            try
            {
                SemanticServiceStatus result = await System.Threading.Tasks.Task
                    .Factory.StartNew(
                        delegate { return DownloadModel(addIn); },
                        CancellationToken.None,
                        System.Threading.Tasks.TaskCreationOptions.LongRunning,
                        System.Threading.Tasks.TaskScheduler.Default);
                _semanticStatus.Text = result.ModelAvailable
                    ? "本地语义模型已就绪。"
                    : "下载模型失败：" + result.ModelError;
            }
            finally
            {
                SetSemanticBusy(false);
                RefreshSemanticStatus();
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
                // The status call below reports the reason.
            }

            return addIn.SemanticClient.GetModelStatus();
        }

        private void UpdateSemanticProgress(SemanticIndexProgress progress)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<SemanticIndexProgress>(
                    UpdateSemanticProgress), progress);
                return;
            }

            _semanticStatus.Text = "索引中：" + progress.Processed
                + " / " + progress.Total
                + "（新增 " + progress.Indexed + "，跳过 " + progress.Skipped
                + "，失败 " + progress.Failed + "）";
        }

        private void SemanticDeleteClicked(object sender, EventArgs e)
        {
            Connect addIn = Connect.Current;
            if (addIn == null || _semanticBusy)
            {
                return;
            }

            if (MessageBox.Show(
                this,
                "将删除本机保存的全部邮件正文副本、片段和向量数据。"
                + "Outlook 中的原始邮件不会被删除。是否继续？",
                "删除本地语义索引",
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Warning) != DialogResult.OK)
            {
                return;
            }

            try
            {
                addIn.SemanticIndexService.ClearIndex();
                _semanticStatus.Text = "本地语义索引已删除。";
            }
            catch (SemanticServiceException ex)
            {
                MessageBox.Show(
                    this,
                    ex.Message,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                RefreshSemanticStatus();
            }
        }

        private void RefreshSemanticStatus()
        {
            Connect addIn = Connect.Current;
            if (addIn == null)
            {
                _semanticStatus.Text = "语义索引状态：加载项正在初始化。";
                return;
            }

            SemanticServiceStatus status = addIn.SemanticClient.GetStatus();
            if (!status.Reachable)
            {
                _semanticStatus.Text =
                    "语义索引状态：本地服务未运行。"
                    + "请先启动 embedding_service 后再建立索引。";
                return;
            }

            _semanticStatus.Text =
                "语义索引状态：已索引 " + status.MailCount + " 封邮件，"
                + status.ChunkCount + " 个片段（"
                + status.EmbeddedChunkCount + " 个已向量化），"
                + "占用 " + FormatSize(status.SizeBytes) + "。\r\n"
                + "模型：" + status.ModelId
                + (status.ModelAvailable
                    ? "（就绪）"
                    : "（未就绪，请点击“下载模型”）")
                + "，设备：" + status.Device
                + (string.IsNullOrWhiteSpace(status.LastSyncAt)
                    ? string.Empty
                    : "\r\n上次同步：" + status.LastSyncAt);
        }

        private void SetSemanticBusy(bool busy)
        {
            _semanticBusy = busy;
            _semanticSyncButton.Enabled = !busy;
            _semanticRebuildButton.Enabled = !busy;
            _semanticDeleteButton.Enabled = !busy;
            _semanticRefreshButton.Enabled = !busy;
            _semanticDownloadButton.Enabled = !busy;
            if (!busy)
            {
                _semanticStopButton.Enabled = false;
            }
        }

        private AppSettings PersistWithoutClosing(AppSettings updated)
        {
            _settingsStore.Save(updated);
            _settings = updated.Copy();
            return updated;
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

        private void SaveClicked(object sender, EventArgs e)
        {
            AppSettings updated;
            if (!TrySaveSettings(out updated))
            {
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private async void RebuildIndexClicked(object sender, EventArgs e)
        {
            AppSettings updated;
            if (_contactIndexService == null
                || !TrySaveSettings(out updated))
            {
                return;
            }

            _rebuildIndex.Enabled = false;
            _indexStatus.Text = "正在扫描 Outlook 本地历史邮件……";
            try
            {
                await _contactIndexService.RebuildAsync(
                    updated,
                    CancellationToken.None,
                    delegate(int completed, int total)
                    {
                        _indexStatus.Text = total == 0
                            ? "正在保存联系人索引……"
                            : "正在分析联系人：" + completed + "/" + total;
                    });
                RefreshIndexStatus();
            }
            catch (Exception ex)
            {
                RefreshIndexStatus();
                MessageBox.Show(
                    this,
                    "重建联系人索引失败：" + ex.Message,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                _rebuildIndex.Enabled = true;
            }
        }

        private bool TrySaveSettings(out AppSettings updated)
        {
            updated = null;
            try
            {
                updated = _settings.Copy();
                updated.ProviderId = string.Empty;
                updated.ApiBaseUrl = (_baseUrl.Text ?? string.Empty).Trim();
                updated.Model = (_model.Text ?? string.Empty).Trim();
                updated.ReasoningEffort = _reasoningEffort.Text.ToLowerInvariant();
                OpenAiCompatibleClient.BuildChatCompletionsUrl(updated.ApiBaseUrl);
                if (string.IsNullOrWhiteSpace(updated.Model))
                {
                    ShowValidation("请填写 Model。");
                    return false;
                }

                bool endpointChanged = !SettingsStore.IsSameApiEndpoint(
                    _settings.ApiBaseUrl,
                    updated.ApiBaseUrl);
                if ((!_settingsStore.HasApiKey(_settings) || endpointChanged)
                    && string.IsNullOrWhiteSpace(_apiKey.Text)
                    && !_clearApiKey.Checked)
                {
                    ShowValidation(endpointChanged
                        ? "更改服务地址后，请重新输入 API Key。"
                        : "请填写 API Key。");
                    return false;
                }
                updated.SummaryLanguage = string.IsNullOrWhiteSpace(_language.Text)
                    ? "简体中文"
                    : _language.Text.Trim();
                updated.MaxEmailCharacters =
                    Decimal.ToInt32(_maximumCharacters.Value);
                updated.IdentityEmailAddresses = _identityEmails.Text.Trim();
                updated.IdentityAliases = _identityAliases.Text.Trim();
                updated.SemanticIndexMonths = (int)_semanticMonths.Value;
                updated.SemanticIndexMaxMails = (int)_semanticMaxMails.Value;
                updated.SemanticIndexIncludeSent = _semanticIncludeSent.Checked;

                if (_clearApiKey.Checked)
                {
                    _settingsStore.SetApiKey(updated, string.Empty);
                }
                else if (!string.IsNullOrWhiteSpace(_apiKey.Text))
                {
                    _settingsStore.SetApiKey(updated, _apiKey.Text);
                }

                _settingsStore.Save(updated);
                _settings = updated.Copy();
                SavedSettings = updated;
                _apiKey.Text = string.Empty;
                _clearApiKey.Checked = false;
                Connect addIn = Connect.Current;
                if (addIn != null)
                {
                    addIn.ApplySettings(updated);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "保存设置失败：" + ex.Message,
                    "AI 邮件助手",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        private void RefreshIndexStatus()
        {
            if (_contactIndexStore == null || !_contactIndexStore.Exists())
            {
                _indexStatus.Text =
                    "联系人索引：未建立\r\nOutlook 归类文件夹：未读取";
                return;
            }

            try
            {
                ContactIndex index = _contactIndexStore.Load();
                DateTime createdAt;
                string displayTime = DateTime.TryParse(
                    index.CreatedAt,
                    out createdAt)
                        ? createdAt.ToString("yyyy-MM-dd HH:mm")
                        : index.CreatedAt;
                _indexStatus.Text =
                    "联系人索引：已建立\r\n联系人数量："
                    + index.Contacts.Count
                    + "\r\n最后更新时间：" + displayTime
                    + "\r\nOutlook 归类文件夹："
                    + index.FolderContext.Count;
            }
            catch
            {
                _indexStatus.Text =
                    "联系人索引：文件无法读取";
            }
        }

        private void ShowValidation(string message)
        {
            MessageBox.Show(
                this,
                message,
                "设置不完整",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        private static TextBox CreateTextBox()
        {
            TextBox textBox = new TextBox();
            textBox.Dock = DockStyle.Top;
            return textBox;
        }

        private static TextBox CreateMultilineTextBox(int height)
        {
            TextBox textBox = CreateTextBox();
            textBox.Multiline = true;
            textBox.ScrollBars = ScrollBars.Vertical;
            textBox.Height = height;
            return textBox;
        }

        private static ComboBox CreateComboBox()
        {
            ComboBox comboBox = new ComboBox();
            comboBox.Dock = DockStyle.Top;
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            return comboBox;
        }

        private static Label CreateNoteLabel()
        {
            Label label = new Label();
            label.AutoSize = true;
            label.MaximumSize = new Size(520, 0);
            label.ForeColor = Color.DimGray;
            return label;
        }

        private static void AddRow(
            TableLayoutPanel panel,
            int row,
            string label,
            Control control)
        {
            Label title = new Label();
            title.Text = label;
            title.AutoSize = true;
            title.Margin = new Padding(3, 7, 8, 3);
            panel.Controls.Add(title, 0, row);
            control.Margin = new Padding(3, 3, 3, 9);
            panel.Controls.Add(control, 1, row);
        }
    }
}
