using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using OutlookAiAssistant.AI;

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
            : this(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAiAssistant"))
        {
        }

        public SettingsStore(string settingsDirectory)
        {
            if (string.IsNullOrWhiteSpace(settingsDirectory))
            {
                throw new ArgumentException(
                    "Settings directory is required.",
                    "settingsDirectory");
            }

            SettingsDirectory = Path.GetFullPath(settingsDirectory);
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
                Dictionary<string, object> rawSettings =
                    _serializer.DeserializeObject(json) as Dictionary<string, object>;
                bool legacySettings = IsLegacySettings(settings);
                if (legacySettings)
                {
                    ApplyLegacyProviderDefaults(
                        settings,
                        HasJsonProperty(rawSettings, "ApiBaseUrl"),
                        HasJsonProperty(rawSettings, "Model"));
                }
                return Normalize(settings, legacySettings);
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
            AppSettings normalized = Normalize(settings, false);
            EnsureApiKeyMatchesEndpoint(normalized);
            string json = _serializer.Serialize(normalized);
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

            EnsureApiKeyMatchesEndpoint(settings);
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
                settings.ApiKeyBaseUrl = string.Empty;
                return;
            }

            string endpoint = OpenAiCompatibleClient.BuildChatCompletionsUrl(
                settings.ApiBaseUrl);
            byte[] plaintext = Encoding.UTF8.GetBytes(apiKey.Trim());
            byte[] cipher = ProtectedData.Protect(
                plaintext,
                Entropy,
                DataProtectionScope.CurrentUser);
            settings.ApiKeyCiphertext = Convert.ToBase64String(cipher);
            settings.ApiKeyBaseUrl = endpoint;
        }

        public static bool IsSameApiEndpoint(string first, string second)
        {
            return string.Equals(
                CanonicalEndpoint(first),
                CanonicalEndpoint(second),
                StringComparison.Ordinal);
        }

        private static AppSettings Normalize(AppSettings settings, bool legacySettings)
        {
            settings = settings ?? new AppSettings();
            if (legacySettings)
            {
                settings.ProviderId = string.Empty;
            }

            settings.ApiBaseUrl = (settings.ApiBaseUrl ?? string.Empty).Trim();
            settings.Model = (settings.Model ?? string.Empty).Trim();
            if (settings.Model.Length == 0)
            {
                throw new InvalidOperationException("模型不能为空。请在设置中输入模型名称。");
            }

            OpenAiCompatibleClient.BuildChatCompletionsUrl(settings.ApiBaseUrl);
            settings.ReasoningEffort = NormalizeReasoningEffort(settings.ReasoningEffort);
            settings.ApiKeyCiphertext = settings.ApiKeyCiphertext ?? string.Empty;
            settings.ApiKeyBaseUrl = (settings.ApiKeyBaseUrl ?? string.Empty).Trim();
            if (legacySettings && settings.ApiKeyCiphertext.Length > 0
                && settings.ApiKeyBaseUrl.Length == 0)
            {
                settings.ApiKeyBaseUrl = OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    settings.ApiBaseUrl);
            }
            settings.SummaryLanguage = string.IsNullOrWhiteSpace(settings.SummaryLanguage)
                ? "简体中文"
                : settings.SummaryLanguage.Trim();
            settings.IdentityEmailAddresses =
                (settings.IdentityEmailAddresses ?? string.Empty).Trim();
            settings.IdentityAliases =
                (settings.IdentityAliases ?? string.Empty).Trim();

            if (settings.MaxEmailCharacters < 5000 || settings.MaxEmailCharacters > 200000)
            {
                settings.MaxEmailCharacters = 40000;
            }

            return settings;
        }

        private static bool IsLegacySettings(AppSettings settings)
        {
            return settings != null
                && (string.Equals(settings.ProviderId, "deepseek", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(settings.ProviderId, "zhipu", StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasJsonProperty(
            Dictionary<string, object> settings,
            string propertyName)
        {
            return settings != null && settings.ContainsKey(propertyName);
        }

        private static void ApplyLegacyProviderDefaults(
            AppSettings settings,
            bool hasBaseUrl,
            bool hasModel)
        {
            if (string.Equals(settings.ProviderId, "zhipu", StringComparison.OrdinalIgnoreCase))
            {
                if (!hasBaseUrl || string.IsNullOrWhiteSpace(settings.ApiBaseUrl))
                {
                    settings.ApiBaseUrl = "https://open.bigmodel.cn/api/paas/v4";
                }

                if (!hasModel || string.IsNullOrWhiteSpace(settings.Model))
                {
                    settings.Model = "glm-5.3-flash";
                }
            }
            else if (string.Equals(settings.ProviderId, "deepseek", StringComparison.OrdinalIgnoreCase))
            {
                if (!hasBaseUrl || string.IsNullOrWhiteSpace(settings.ApiBaseUrl))
                {
                    settings.ApiBaseUrl = "https://api.deepseek.com";
                }

                if (!hasModel || string.IsNullOrWhiteSpace(settings.Model))
                {
                    settings.Model = "deepseek-v4-flash";
                }
            }
        }

        private static string NormalizeReasoningEffort(string value)
        {
            value = (value ?? string.Empty).Trim().ToLowerInvariant();
            return value == "medium" || value == "max" || value == "none"
                ? value
                : "none";
        }

        private static void EnsureApiKeyMatchesEndpoint(AppSettings settings)
        {
            if (settings == null || string.IsNullOrWhiteSpace(settings.ApiKeyCiphertext))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(settings.ApiKeyBaseUrl)
                || !IsSameApiEndpoint(settings.ApiKeyBaseUrl, settings.ApiBaseUrl))
            {
                throw new InvalidOperationException(
                    "API 地址已修改，请重新输入 API Key 后再保存或使用。");
            }
        }

        private static string CanonicalEndpoint(string baseUrl)
        {
            string endpoint = OpenAiCompatibleClient.BuildChatCompletionsUrl(baseUrl);
            Uri uri = new Uri(endpoint, UriKind.Absolute);
            string host = uri.Host.ToLowerInvariant();
            bool defaultPort = (uri.Scheme == Uri.UriSchemeHttps && uri.Port == 443)
                || (uri.Scheme == Uri.UriSchemeHttp && uri.Port == 80);
            string port = defaultPort ? string.Empty : ":" + uri.Port;
            string path = uri.AbsolutePath.TrimEnd('/');
            return uri.Scheme.ToLowerInvariant() + "://" + host + port + path;
        }
    }
}

