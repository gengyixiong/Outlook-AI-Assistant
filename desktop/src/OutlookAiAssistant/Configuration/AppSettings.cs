namespace OutlookAiAssistant.Configuration
{
    /// <summary>
    /// User-editable configuration. ApiKeyCiphertext contains only a DPAPI
    /// encrypted value and is safe to persist in the local settings file.
    /// </summary>
    public sealed class AppSettings
    {
        // Retained only to migrate settings written by older releases.
        public string ProviderId { get; set; }
        public string ApiBaseUrl { get; set; }
        public string Model { get; set; }
        public string ReasoningEffort { get; set; }
        public string ApiKeyCiphertext { get; set; }
        // Chat Completions endpoint to which the saved key belongs.
        public string ApiKeyBaseUrl { get; set; }
        public string SummaryLanguage { get; set; }
        public int MaxEmailCharacters { get; set; }
        public string IdentityEmailAddresses { get; set; }
        public string IdentityAliases { get; set; }

        public AppSettings()
        {
            ProviderId = string.Empty;
            ApiBaseUrl = "https://api.openai.com/v1";
            Model = "gpt-5.6-luna";
            ReasoningEffort = "none";
            ApiKeyCiphertext = string.Empty;
            ApiKeyBaseUrl = string.Empty;
            SummaryLanguage = "简体中文";
            MaxEmailCharacters = 40000;
            IdentityEmailAddresses = string.Empty;
            IdentityAliases = string.Empty;
        }

        public AppSettings Copy()
        {
            return new AppSettings
            {
                ProviderId = ProviderId,
                ApiBaseUrl = ApiBaseUrl,
                Model = Model,
                ReasoningEffort = ReasoningEffort,
                ApiKeyCiphertext = ApiKeyCiphertext,
                ApiKeyBaseUrl = ApiKeyBaseUrl,
                SummaryLanguage = SummaryLanguage,
                MaxEmailCharacters = MaxEmailCharacters,
                IdentityEmailAddresses = IdentityEmailAddresses,
                IdentityAliases = IdentityAliases
            };
        }
    }
}

