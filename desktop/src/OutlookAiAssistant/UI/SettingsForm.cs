using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;
using OutlookAiAssistant.AddIn;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.EntityIndex;

namespace OutlookAiAssistant.UI
{
    /// <summary>
    /// User settings for the two fixed-cost Flash providers, summary identity
    /// and explicitly initiated Outlook contact-index rebuilds.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly SettingsStore _settingsStore;
        private AppSettings _settings;
        private readonly IList<AiProviderPreset> _presets;
        private readonly ContactIndexStore _contactIndexStore;
        private readonly ContactIndexService _contactIndexService;
        private ComboBox _provider;
        private TextBox _apiKey;
        private CheckBox _showApiKey;
        private CheckBox _clearApiKey;
        private ComboBox _language;
        private NumericUpDown _maximumCharacters;
        private TextBox _identityEmails;
        private TextBox _identityAliases;
        private Label _providerHelp;
        private Label _indexStatus;
        private Button _rebuildIndex;
        private bool _loading;

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
            _presets = AiProviderPreset.CreateDefaults();
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
            root.RowCount = 11;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            scroll.Controls.Add(root);

            _provider = CreateComboBox();
            _provider.DisplayMember = "DisplayName";
            for (int index = 0; index < _presets.Count; index++)
            {
                _provider.Items.Add(_presets[index]);
            }
            _provider.SelectedIndexChanged += ProviderSelectedIndexChanged;
            AddRow(root, 0, "AI Provider", _provider);

            _providerHelp = CreateNoteLabel();
            root.Controls.Add(_providerHelp, 1, 1);

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
            AddRow(root, 2, "API Key", apiKeyPanel);

            Label keyNote = CreateNoteLabel();
            keyNote.Text =
                "留空表示继续使用已保存的密钥。密钥使用 Windows DPAPI 加密，"
                + "不会在窗口中回显。模型由 Provider 固定，不能切换到 Pro。";
            root.Controls.Add(keyNote, 1, 3);

            _language = CreateComboBox();
            _language.Items.AddRange(
                new object[] { "简体中文", "English", "跟随邮件语言" });
            AddRow(root, 4, "摘要语言", _language);

            _maximumCharacters = new NumericUpDown();
            _maximumCharacters.Dock = DockStyle.Top;
            _maximumCharacters.Minimum = 5000;
            _maximumCharacters.Maximum = 200000;
            _maximumCharacters.Increment = 5000;
            _maximumCharacters.ThousandsSeparator = true;
            AddRow(root, 5, "正文字符上限", _maximumCharacters);

            _identityEmails = CreateMultilineTextBox(64);
            AddRow(root, 6, "我的邮箱地址", _identityEmails);

            _identityAliases = CreateMultilineTextBox(82);
            AddRow(root, 7, "别人对我的称呼", _identityAliases);

            _indexStatus = new Label();
            _indexStatus.AutoSize = true;
            _indexStatus.MaximumSize = new Size(520, 0);
            AddRow(root, 8, "联系人索引", _indexStatus);

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
            AddRow(root, 9, string.Empty, rebuildPanel);

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
            root.Controls.Add(buttons, 0, 10);
            root.SetColumnSpan(buttons, 2);
            AcceptButton = save;
            CancelButton = cancel;
        }

        private void LoadSettings()
        {
            _loading = true;
            try
            {
                int selectedIndex = 0;
                for (int index = 0; index < _presets.Count; index++)
                {
                    if (string.Equals(
                        _presets[index].Id,
                        _settings.ProviderId,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        selectedIndex = index;
                        break;
                    }
                }

                if (_provider.Items.Count > 0)
                {
                    _provider.SelectedIndex = Math.Min(
                        selectedIndex,
                        _provider.Items.Count - 1);
                }
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
                UpdateProviderHelp();
            }
            finally
            {
                _loading = false;
            }
        }

        private void ProviderSelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_loading)
            {
                UpdateProviderHelp();
            }
        }

        private void UpdateProviderHelp()
        {
            AiProviderPreset preset = _provider.SelectedItem as AiProviderPreset;
            _providerHelp.Text = preset == null
                ? string.Empty
                : preset.HelpText + "\r\n固定模型：" + preset.DefaultModel;
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
            AiProviderPreset preset = _provider.SelectedItem as AiProviderPreset;
            if (preset == null)
            {
                ShowValidation("请选择 AI Provider。");
                return false;
            }

            bool hasStoredKey = _settingsStore.HasApiKey(_settings);
            bool providerChanged = !string.Equals(
                preset.Id,
                _settings.ProviderId,
                StringComparison.OrdinalIgnoreCase);
            if ((!hasStoredKey || providerChanged)
                && string.IsNullOrWhiteSpace(_apiKey.Text)
                && !_clearApiKey.Checked)
            {
                ShowValidation(providerChanged
                    ? "切换 Provider 后，请填写该 Provider 的 API Key。"
                    : "请填写 API Key。");
                return false;
            }

            try
            {
                updated = _settings.Copy();
                updated.ProviderId = preset.Id;
                updated.ApiBaseUrl = preset.BaseUrl;
                updated.Model = preset.DefaultModel;
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
