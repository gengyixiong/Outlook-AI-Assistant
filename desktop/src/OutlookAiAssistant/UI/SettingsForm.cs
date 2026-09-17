using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.EntityIndex;

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
            root.RowCount = 13;
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
            root.Controls.Add(buttons, 0, 12);
            root.SetColumnSpan(buttons, 2);
            AcceptButton = save;
            CancelButton = cancel;
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
