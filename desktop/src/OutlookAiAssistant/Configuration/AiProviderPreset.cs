using System.Collections.Generic;

namespace OutlookAiAssistant.Configuration
{
    /// <summary>
    /// A UI preset only. Users can always override the endpoint and model.
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
                    "DeepSeek",
                    "https://api.deepseek.com",
                    "deepseek-v4-flash",
                    "适合优先考虑成本与速度的摘要和查询解析。"),
                new AiProviderPreset(
                    "openai",
                    "OpenAI",
                    "https://api.openai.com/v1",
                    "gpt-5.6-sol",
                    "使用 OpenAI Chat Completions 兼容接口。"),
                new AiProviderPreset(
                    "doubao",
                    "豆包 / 火山方舟",
                    "https://ark.cn-beijing.volces.com/api/v3",
                    string.Empty,
                    "模型字段请填写火山方舟控制台提供的模型或推理接入点 ID。"),
                new AiProviderPreset(
                    "custom",
                    "自定义 OpenAI 兼容接口",
                    string.Empty,
                    string.Empty,
                    "适用于其他提供 /chat/completions 的兼容服务。")
            };
        }
    }
}

