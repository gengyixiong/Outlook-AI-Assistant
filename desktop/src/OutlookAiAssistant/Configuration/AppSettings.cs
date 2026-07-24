namespace OutlookAiAssistant.Configuration
{
    /// <summary>
    /// User-editable configuration. ApiKeyCiphertext contains only a DPAPI
    /// encrypted value and is safe to persist in the local settings file.
    /// </summary>
    public sealed class AppSettings
    {
        public string ProviderId { get; set; }
        public string ApiBaseUrl { get; set; }
        public string Model { get; set; }
        public string ApiKeyCiphertext { get; set; }
        public string SummaryLanguage { get; set; }
        public int MaxEmailCharacters { get; set; }

        public AppSettings()
        {
            ProviderId = "deepseek";
            ApiBaseUrl = "https://api.deepseek.com";
            Model = "deepseek-v4-flash";
            ApiKeyCiphertext = string.Empty;
            SummaryLanguage = "简体中文";
            MaxEmailCharacters = 40000;
        }

        public AppSettings Copy()
        {
            return new AppSettings
            {
                ProviderId = ProviderId,
                ApiBaseUrl = ApiBaseUrl,
                Model = Model,
                ApiKeyCiphertext = ApiKeyCiphertext,
                SummaryLanguage = SummaryLanguage,
                MaxEmailCharacters = MaxEmailCharacters
            };
        }
    }
}

