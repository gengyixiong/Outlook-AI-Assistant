using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookAiAssistant.Configuration
{
    /// <summary>
    /// Stores non-secret settings as JSON and encrypts the API key with Windows
    /// DPAPI for the current Windows user.
    /// </summary>
    public sealed class SettingsStore
    {
        private static readonly byte[] Entropy =
            Encoding.UTF8.GetBytes("OutlookAiAssistant.Settings.v1");

        private readonly JavaScriptSerializer _serializer;

        public string SettingsDirectory { get; private set; }
        public string SettingsPath { get; private set; }

        public SettingsStore()
        {
            SettingsDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAiAssistant");
            SettingsPath = Path.Combine(SettingsDirectory, "settings.json");
            _serializer = new JavaScriptSerializer();
        }

        public AppSettings Load()
        {
            if (!File.Exists(SettingsPath))
            {
                return new AppSettings();
            }

            try
            {
                string json = File.ReadAllText(SettingsPath, Encoding.UTF8);
                AppSettings settings = _serializer.Deserialize<AppSettings>(json);
                return Normalize(settings);
            }
            catch
            {
                // A corrupt settings file must not prevent Outlook from loading.
                return new AppSettings();
            }
        }

        public void Save(AppSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            Directory.CreateDirectory(SettingsDirectory);
            string json = _serializer.Serialize(Normalize(settings));
            string temporaryPath = SettingsPath + ".tmp";
            File.WriteAllText(temporaryPath, json, new UTF8Encoding(false));

            if (File.Exists(SettingsPath))
            {
                File.Replace(temporaryPath, SettingsPath, null);
            }
            else
            {
                File.Move(temporaryPath, SettingsPath);
            }
        }

        public bool HasApiKey(AppSettings settings)
        {
            return settings != null
                && !string.IsNullOrWhiteSpace(settings.ApiKeyCiphertext);
        }

        public string ReadApiKey(AppSettings settings)
        {
            if (!HasApiKey(settings))
            {
                return string.Empty;
            }

            byte[] cipher = Convert.FromBase64String(settings.ApiKeyCiphertext);
            byte[] plaintext = ProtectedData.Unprotect(
                cipher,
                Entropy,
                DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plaintext);
        }

        public void SetApiKey(AppSettings settings, string apiKey)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                settings.ApiKeyCiphertext = string.Empty;
                return;
            }

            byte[] plaintext = Encoding.UTF8.GetBytes(apiKey.Trim());
            byte[] cipher = ProtectedData.Protect(
                plaintext,
                Entropy,
                DataProtectionScope.CurrentUser);
            settings.ApiKeyCiphertext = Convert.ToBase64String(cipher);
        }

        private static AppSettings Normalize(AppSettings settings)
        {
            settings = settings ?? new AppSettings();
            settings.ProviderId = string.IsNullOrWhiteSpace(settings.ProviderId)
                ? "deepseek"
                : settings.ProviderId.Trim();
            settings.ApiBaseUrl = (settings.ApiBaseUrl ?? string.Empty).Trim();
            settings.Model = (settings.Model ?? string.Empty).Trim();
            settings.ApiKeyCiphertext = settings.ApiKeyCiphertext ?? string.Empty;
            settings.SummaryLanguage = string.IsNullOrWhiteSpace(settings.SummaryLanguage)
                ? "简体中文"
                : settings.SummaryLanguage.Trim();

            if (settings.MaxEmailCharacters < 5000 || settings.MaxEmailCharacters > 200000)
            {
                settings.MaxEmailCharacters = 40000;
            }

            return settings;
        }
    }
}

