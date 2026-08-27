using System;
using System.Threading;
using System.Threading.Tasks;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Summary
{
    public sealed class SummaryService
    {
        private readonly OpenAiCompatibleClient _client;
        private readonly SettingsStore _settingsStore;
        private readonly ConversationPromptBuilder _promptBuilder;

        public SummaryService(
            OpenAiCompatibleClient client,
            SettingsStore settingsStore)
        {
            _client = client;
            _settingsStore = settingsStore;
            _promptBuilder = new ConversationPromptBuilder();
        }

        public Task<string> SummarizeAsync(
            ConversationSnapshot conversation,
            AppSettings settings,
            CancellationToken cancellationToken)
        {
            if (conversation == null)
            {
                throw new ArgumentNullException("conversation");
            }

            string apiKey = _settingsStore.ReadApiKey(settings);
            string systemPrompt = _promptBuilder.BuildSystemPrompt(
                settings.SummaryLanguage);
            string userPrompt = _promptBuilder.BuildUserPrompt(
                conversation,
                settings.MaxEmailCharacters,
                settings.IdentityEmailAddresses,
                settings.IdentityAliases);
            return _client.CompleteAsync(
                settings,
                apiKey,
                systemPrompt,
                userPrompt,
                false,
                cancellationToken);
        }
    }
}
