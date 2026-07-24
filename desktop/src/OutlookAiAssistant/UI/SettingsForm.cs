using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using OutlookAiAssistant.Configuration;

namespace OutlookAiAssistant.UI
{
    /// <summary>
    /// Provider configuration dialog. The stored API key is never placed back
    /// into the text box, which prevents accidental disclosure.
    /// </summary>
    public sealed class SettingsForm : Form
    {
        private readonly SettingsStore _settingsStore;
        private readonly AppSettings _originalSettings;
        private readonly IList<AiProviderPreset> _presets;
        private ComboBox _provider;
        private TextBox _baseUrl;
        private TextBox _model;
        private TextBox _apiKey;
        private CheckBox _showApiKey;
        private CheckBox _clearApiKey;
        private ComboBox _language;
        private NumericUpDown _maximumCharacters;
        private Label _providerHelp;
        private bool _loading;

        public AppSettings SavedSettings { get; private set; }

        public SettingsForm(
            SettingsStore settingsStore,
            AppSettings settings)
        {
            _settingsStore = settingsStore;
            _originalSettings = settings.Copy();
            _presets = AiProviderPreset.CreateDefaults();
            Text = "AI 邮件助手设置";
            StartPosition = FormStartPosition.CenterParent;
            MinimumSize = new Size(640, 520);
            Size = new Size(700, 600);
            Font = new Font("Segoe UI", 9F);
            ShowIcon = false;
            BuildInterface();
            LoadSettings();
        }

        private void BuildInterface()
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.Padding = new Padding(18);
            root.ColumnCount = 2;
            root.RowCount = 9;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            Controls.Add(root);

            _provider = CreateComboBox();
            _provider.DisplayMember = "DisplayName";
            // Populate the items directly. Assigning DataSource before this
            // control has a BindingContext can leave Items empty until the
            // form handle is created, while LoadSettings runs immediately.
            for (int index = 0; index < _presets.Count; index++)
            {
                _provider.Items.Add(_presets[index]);
            }
            _provider.SelectedIndexChanged += ProviderSelectedIndexChanged;
            AddRow(root, 0, "API 提供商", _provider);

            _providerHelp = new Label();
            _providerHelp.AutoSize = true;
            _providerHelp.MaximumSize = new Size(470, 0);
            _providerHelp.ForeColor = Color.DimGray;
            root.Controls.Add(_providerHelp, 1, 1);

            _baseUrl = CreateTextBox();
            AddRow(root, 2, "API 地址", _baseUrl);

            _model = CreateTextBox();
            AddRow(root, 3, "模型 / 接入点 ID", _model);

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
            AddRow(root, 4, "API Key", apiKeyPanel);

            Label keyNote = new Label();
            keyNote.AutoSize = true;
            keyNote.MaximumSize = new Size(470, 0);
            keyNote.ForeColor = Color.DimGray;
            keyNote.Text =
                "留空表示继续使用已保存的密钥。密钥使用 Windows DPAPI 加密，"
                + "只能由当前 Windows 用户解密。";
            root.Controls.Add(keyNote, 1, 5);

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
            AddRow(root, 7, "本次发送正文字符上限", _maximumCharacters);

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
            root.Controls.Add(buttons, 0, 8);
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
                        _originalSettings.ProviderId,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        selectedIndex = index;
                        break;
                    }
                }

                // Keep the settings dialog usable even if a future build is
                // accidentally shipped without any provider presets.
                if (_provider.Items.Count > 0)
                {
                    _provider.SelectedIndex = Math.Min(
                        selectedIndex,
                        _provider.Items.Count - 1);
                }
                _baseUrl.Text = _originalSettings.ApiBaseUrl;
                _model.Text = _originalSettings.Model;
                _apiKey.Text = string.Empty;
                _clearApiKey.Checked = false;
                _language.Text = _originalSettings.SummaryLanguage;
                _maximumCharacters.Value = Math.Max(
                    _maximumCharacters.Minimum,
                    Math.Min(
                        _maximumCharacters.Maximum,
                        _originalSettings.MaxEmailCharacters));
                UpdateProviderHelp();
            }
            finally
            {
                _loading = false;
            }
        }

        private void ProviderSelectedIndexChanged(object sender, EventArgs e)
        {
            AiProviderPreset preset = _provider.SelectedItem as AiProviderPreset;
            if (preset == null)
            {
                return;
            }

            UpdateProviderHelp();
            if (_loading)
            {
                return;
            }

            _baseUrl.Text = preset.BaseUrl;
            _model.Text = preset.DefaultModel;
        }

        private void UpdateProviderHelp()
        {
            AiProviderPreset preset = _provider.SelectedItem as AiProviderPreset;
            _providerHelp.Text = preset == null ? string.Empty : preset.HelpText;
        }

        private void SaveClicked(object sender, EventArgs e)
        {
            AiProviderPreset preset = _provider.SelectedItem as AiProviderPreset;
            if (preset == null
                || string.IsNullOrWhiteSpace(_baseUrl.Text)
                || string.IsNullOrWhiteSpace(_model.Text))
            {
                MessageBox.Show(
                    this,
                    "请选择提供商，并填写 API 地址和模型 / 接入点 ID。",
                    "设置不完整",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            bool hasStoredKey = _settingsStore.HasApiKey(_originalSettings);
            if (!hasStoredKey
                && string.IsNullOrWhiteSpace(_apiKey.Text)
                && !_clearApiKey.Checked)
            {
                MessageBox.Show(
                    this,
                    "请填写 API Key。",
                    "设置不完整",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            try
            {
                AppSettings updated = _originalSettings.Copy();
                updated.ProviderId = preset.Id;
                updated.ApiBaseUrl = _baseUrl.Text.Trim();
                updated.Model = _model.Text.Trim();
                updated.SummaryLanguage = string.IsNullOrWhiteSpace(_language.Text)
                    ? "简体中文"
                    : _language.Text.Trim();
                updated.MaxEmailCharacters =
                    Decimal.ToInt32(_maximumCharacters.Value);

                if (_clearApiKey.Checked)
                {
                    _settingsStore.SetApiKey(updated, string.Empty);
                }
                else if (!string.IsNullOrWhiteSpace(_apiKey.Text))
                {
                    _settingsStore.SetApiKey(updated, _apiKey.Text);
                }

                _settingsStore.Save(updated);
                SavedSettings = updated;
                DialogResult = DialogResult.OK;
                Close();
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

        private static TextBox CreateTextBox()
        {
            TextBox textBox = new TextBox();
            textBox.Dock = DockStyle.Top;
            return textBox;
        }

        private static ComboBox CreateComboBox()
        {
            ComboBox comboBox = new ComboBox();
            comboBox.Dock = DockStyle.Top;
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            return comboBox;
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
            control.Margin = new Padding(3, 3, 3, 8);
            panel.Controls.Add(control, 1, row);
        }
    }
}
