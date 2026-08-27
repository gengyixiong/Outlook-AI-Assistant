using System.Collections.Generic;
using System;

namespace OutlookAiAssistant.Configuration
{
    /// <summary>
    /// Allow-listed provider configuration. Endpoint and model are intentionally
    /// fixed so the UI cannot drift to an unsupported or costly model.
    /// </summary>
    public sealed class AiProviderPreset
    {
        public string Id { get; private set; }
        public string DisplayName { get; private set; }
        public string BaseUrl { get; private set; }
        public string DefaultModel { get; private set; }
        public string HelpText { get; private set; }

        public AiProviderPreset(
            string id,
            string displayName,
            string baseUrl,
            string defaultModel,
            string helpText)
        {
            Id = id;
            DisplayName = displayName;
            BaseUrl = baseUrl;
            DefaultModel = defaultModel;
            HelpText = helpText;
        }

        public override string ToString()
        {
            return DisplayName;
        }

        public static IList<AiProviderPreset> CreateDefaults()
        {
            return new List<AiProviderPreset>
            {
                new AiProviderPreset(
                    "deepseek",
                    "DeepSeek Flash",
                    "https://api.deepseek.com",
                    "deepseek-v4-flash",
                    "固定使用 DeepSeek Flash 以控制 API 成本。"),
                new AiProviderPreset(
                    "zhipu",
                    "GLM-5.3 Flash",
                    "https://open.bigmodel.cn/api/paas/v4",
                    "glm-5.3-flash",
                    "固定使用智谱 GLM-5.3 Flash 以控制 API 成本。")
            };
        }

        public static AiProviderPreset Find(string providerId)
        {
            IList<AiProviderPreset> presets = CreateDefaults();
            for (int index = 0; index < presets.Count; index++)
            {
                if (string.Equals(
                    presets[index].Id,
                    providerId,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return presets[index];
                }
            }

            return presets[0];
        }
    }
}

