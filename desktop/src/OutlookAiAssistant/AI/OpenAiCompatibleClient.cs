using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using OutlookAiAssistant.Configuration;

namespace OutlookAiAssistant.AI
{
    /// <summary>
    /// Minimal provider-neutral client for OpenAI-compatible Chat Completions.
    /// It intentionally sends no telemetry and never logs prompts or API keys.
    /// </summary>
    public sealed class OpenAiCompatibleClient
    {
        private readonly JavaScriptSerializer _serializer;

        public OpenAiCompatibleClient()
        {
            _serializer = new JavaScriptSerializer();
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
        }

        public Task<string> CompleteAsync(
            AppSettings settings,
            string apiKey,
            string systemPrompt,
            string userPrompt,
            bool requestJson,
            CancellationToken cancellationToken)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            return Task.Factory.StartNew(
                delegate
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    try
                    {
                        return CompleteOnce(
                            settings,
                            apiKey,
                            systemPrompt,
                            userPrompt,
                            requestJson);
                    }
                    catch (AiHttpException ex)
                    {
                        // Some compatible services do not implement response_format.
                        // Retry only that compatibility failure without JSON mode.
                        if (requestJson && ex.StatusCode == HttpStatusCode.BadRequest)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            return CompleteOnce(
                                settings,
                                apiKey,
                                systemPrompt,
                                userPrompt,
                                false);
                        }

                        throw;
                    }
                },
                cancellationToken);
        }

        public static string BuildChatCompletionsUrl(string baseUrl)
        {
            string value = (baseUrl ?? string.Empty).Trim().TrimEnd('/');
            if (value.Length == 0)
            {
                throw new InvalidOperationException("API 地址不能为空。");
            }

            Uri uri;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri))
            {
                throw new InvalidOperationException("API 地址格式无效。");
            }

            bool isLocalHttp = uri.Scheme == Uri.UriSchemeHttp
                && (uri.Host == "localhost" || uri.Host == "127.0.0.1");
            if (uri.Scheme != Uri.UriSchemeHttps && !isLocalHttp)
            {
                throw new InvalidOperationException(
                    "远程 API 必须使用 HTTPS；只有 localhost 可以使用 HTTP。");
            }

            if (value.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }

            return value + "/chat/completions";
        }

        private string CompleteOnce(
            AppSettings settings,
            string apiKey,
            string systemPrompt,
            string userPrompt,
            bool requestJson)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                throw new InvalidOperationException("请先在设置中填写 API Key。");
            }

            if (string.IsNullOrWhiteSpace(settings.Model))
            {
                throw new InvalidOperationException("请先在设置中填写模型名称。");
            }

            string endpoint = BuildChatCompletionsUrl(settings.ApiBaseUrl);
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model.Trim();
            payload["stream"] = false;
            payload["messages"] = new object[]
            {
                new Dictionary<string, object>
                {
                    { "role", "system" },
                    { "content", systemPrompt ?? string.Empty }
                },
                new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", userPrompt ?? string.Empty }
                }
            };

            if (requestJson)
            {
                payload["response_format"] = new Dictionary<string, object>
                {
                    { "type", "json_object" }
                };
            }

            // DeepSeek V4 enables thinking by default. These two short tasks do
            // not require chain-of-thought, so disable it to reduce latency/cost.
            if (string.Equals(
                settings.ProviderId,
                "deepseek",
                StringComparison.OrdinalIgnoreCase))
            {
                payload["thinking"] = new Dictionary<string, object>
                {
                    { "type", "disabled" }
                };
            }

            byte[] body = Encoding.UTF8.GetBytes(_serializer.Serialize(payload));
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST";
            request.ContentType = "application/json; charset=utf-8";
            request.Accept = "application/json";
            request.Headers[HttpRequestHeader.Authorization] = "Bearer " + apiKey.Trim();
            request.ContentLength = body.Length;
            request.Timeout = 120000;
            request.ReadWriteTimeout = 120000;
            request.UserAgent = "OutlookAiAssistant/0.1";

            try
            {
                using (Stream requestStream = request.GetRequestStream())
                {
                    requestStream.Write(body, 0, body.Length);
                }

                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                using (Stream responseStream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(responseStream, Encoding.UTF8))
                {
                    string json = reader.ReadToEnd();
                    return ExtractText(json);
                }
            }
            catch (WebException ex)
            {
                HttpWebResponse response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    HttpStatusCode statusCode = response.StatusCode;
                    string statusDescription = response.StatusDescription;
                    response.Close();
                    throw new AiHttpException(
                        statusCode,
                        "AI 服务请求失败：" + (int)statusCode + " "
                            + statusDescription,
                        ex);
                }

                throw new InvalidOperationException(
                    "无法连接 AI 服务，请检查网络、API 地址和代理设置。",
                    ex);
            }
        }

        private string ExtractText(string json)
        {
            object rootObject = _serializer.DeserializeObject(json);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null || !root.ContainsKey("choices"))
            {
                throw new InvalidOperationException("AI 服务返回了无法识别的数据。");
            }

            object[] choices = root["choices"] as object[];
            if (choices == null || choices.Length == 0)
            {
                throw new InvalidOperationException("AI 服务没有返回结果。");
            }

            Dictionary<string, object> choice = choices[0] as Dictionary<string, object>;
            object rawMessage;
            Dictionary<string, object> message = choice != null
                && choice.TryGetValue("message", out rawMessage)
                ? rawMessage as Dictionary<string, object>
                : null;
            object content;
            if (message == null
                || !message.TryGetValue("content", out content)
                || content == null)
            {
                throw new InvalidOperationException("AI 服务返回结果中没有文本内容。");
            }

            string text = Convert.ToString(content);
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("AI 服务返回了空内容。");
            }

            return text.Trim();
        }
    }

    public sealed class AiHttpException : Exception
    {
        public HttpStatusCode StatusCode { get; private set; }

        public AiHttpException(
            HttpStatusCode statusCode,
            string message,
            Exception innerException)
            : base(message, innerException)
        {
            StatusCode = statusCode;
        }
    }
}
