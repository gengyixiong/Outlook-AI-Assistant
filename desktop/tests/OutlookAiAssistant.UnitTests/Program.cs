using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Models;
using OutlookAiAssistant.SemanticSearch;
using OutlookAiAssistant.Summary;
using OutlookAiAssistant.UI;

namespace OutlookAiAssistant.UnitTests
{
    internal static class Program
    {
        private static int _passed;
        private static int _failed;

        [STAThread]
        private static int Main()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Run("Builds compatible endpoint URLs", BuildsCompatibleEndpointUrls);
            Run("Builds settings dialog before first display", BuildsSettingsDialog);
            Run("Builds a visible semantic result card", BuildsVisibleResultCard);
            Run("Trims quoted history from stored messages", TrimsQuotedHistory);
            Run("Keeps ordinary From lines", KeepsOrdinaryFromLines);
            Run("Builds separate history and current prompt", BuildsConversationPrompt);
            Run("Uses generic configuration defaults", UsesGenericConfigurationDefaults);
            Run("Round-trips generic configuration and efforts", RoundTripsGenericConfiguration);
            Run("Migrates legacy provider configuration", MigratesLegacyProviderConfiguration);
            Run("Protects API keys when endpoint changes", ProtectsApiKeysAcrossEndpointChanges);
            Run("Sends reasoning effort without thinking", SendsReasoningEffortWithoutThinking);
            Run("Ignores legacy Windows folder setting", IgnoresLegacyFolderSetting);
            Run("Builds executive brief identity prompt", BuildsExecutiveBriefPrompt);
            Run("Shows editable generic AI settings", ShowsEditableGenericSettings);
            Run("Serializes semantic index payload", SerializesSemanticIndexPayload);
            Run("Falls back to loopback semantic service url", FallsBackToLoopbackServiceUrl);
            Run("Rejects non-loopback semantic service url", RejectsNonLoopbackServiceUrl);
            Run("Reports missing semantic service clearly", ReportsMissingSemanticService);
            Run("Parses semantic search response", ParsesSemanticSearchResponse);
            Run("Parses semantic health status", ParsesSemanticHealthStatus);
            Run("Preserves semantic index settings", PreservesSemanticIndexSettings);
            Run("Clamps invalid semantic index settings", ClampsInvalidSemanticIndexSettings);
            Run("Treats 503 as model unavailable", Treats503AsModelUnavailable);
            Run("Reports model status from the local service", ReportsModelStatus);

            Console.WriteLine(
                "Tests complete. Passed: {0}; Failed: {1}",
                _passed,
                _failed);
            return _failed == 0 ? 0 : 1;
        }











        private static void BuildsCompatibleEndpointUrls()
        {
            Equal(
                "https://api.deepseek.com/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://api.deepseek.com/"));
            Equal(
                "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://open.bigmodel.cn/api/paas/v4/chat/completions"));
            Equal(
                "http://example.com/v1/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "http://example.com/v1"));
            Equal(
                "http://localhost:1234/v1/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "http://localhost:1234/v1/"));
            Throws<InvalidOperationException>(delegate
            {
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://example.com/v1?x=1");
            });
            Throws<InvalidOperationException>(delegate
            {
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://user:pass@example.com/v1");
            });
        }

        private static void BuildsSettingsDialog()
        {
            using (SettingsForm form = new SettingsForm(
                new SettingsStore(),
                new AppSettings()))
            {
                Equal("AI 邮件助手设置", form.Text);
            }
        }

        // A previous layout used nested Dock=Top panels that collapsed to an
        // 18-pixel-wide sliver inside the scrolling results panel, so every
        // search result was invisible and unclickable. This pins the card to a
        // usable size.
        private static void BuildsVisibleResultCard()
        {
            SemanticSearchResult result = new SemanticSearchResult();
            result.EntryId = "E1";
            result.Subject = "Re: Curtain wall edit - can we also edit the "
                + "Panels individually and export them to a spreadsheet";
            result.SenderName = "India Client";
            result.SenderEmail = "india@example.com";
            result.ReceivedAt = "2026-06-15T09:00:00Z";
            result.FolderPath = "Inbox";
            result.Snippet = "Regarding the curtain wall edit, can we also "
                + "edit the Panels individually?";

            Control card = AssistantPaneControl.BuildResultCardControl(
                result, 1, new Font("Segoe UI", 9F),
                delegate { }, delegate { });

            using (Form host = new Form())
            {
                host.Width = 410;
                host.Height = 600;
                FlowLayoutPanel panel = new FlowLayoutPanel();
                panel.Dock = DockStyle.Fill;
                panel.FlowDirection = FlowDirection.TopDown;
                panel.WrapContents = false;
                panel.AutoScroll = true;
                panel.Controls.Add(card);
                host.Controls.Add(panel);
                host.Show();
                Application.DoEvents();
                panel.PerformLayout();
                card.PerformLayout();
                host.PerformLayout();
                Application.DoEvents();

                True(card.Width >= 200,
                    "result card must be wide enough to read, got " + card.Width);
                True(card.Height >= 80,
                    "result card must be tall enough to read, got " + card.Height);
                True(card.Width <= panel.Width + 40,
                    "result card must fit the results panel, got " + card.Width
                    + " in " + panel.Width);

                Button open = FindButton(card, "打开邮件");
                Button outlook = FindButton(card, "Outlook 搜索");
                True(open != null, "card must expose an open button");
                True(outlook != null, "card must expose an Outlook search button");
                // The buttons must be reachable (inside the card bounds).
                True(open.Width > 0 && open.Height > 0, "open button must render");
                host.Close();
            }
        }

        private static Button FindButton(Control root, string text)
        {
            foreach (Control child in root.Controls)
            {
                Button button = child as Button;
                if (button != null && button.Text == text)
                {
                    return button;
                }

                Button nested = FindButton(child, text);
                if (nested != null)
                {
                    return nested;
                }
            }

            return null;
        }

        private static void TrimsQuotedHistory()
        {
            string body =
                "New reply text.\r\n\r\n"
                + "From: Previous Sender\r\n"
                + "Sent: Thursday\r\n"
                + "To: Recipient\r\n"
                + "Subject: Earlier message\r\n"
                + "Old quoted text.";
            Equal(
                "New reply text.",
                EmailBodyCleaner.RemoveQuotedHistory(body));
        }

        private static void KeepsOrdinaryFromLines()
        {
            const string body =
                "From: scratch, the team rebuilt the test environment.";
            Equal(body, EmailBodyCleaner.RemoveQuotedHistory(body));
        }

        private static void BuildsConversationPrompt()
        {
            EmailSnapshot current = new EmailSnapshot();
            current.Subject = "Current subject";
            current.Body =
                "CURRENT_UNIQUE\r\n"
                + "From: External history\r\n"
                + "Sent: Yesterday\r\n"
                + "To: Team\r\n"
                + "Subject: Prior context\r\n"
                + "EMBEDDED_CONTEXT";
            current.CurrentUserName = "Current User";
            current.CurrentUserEmail = "current@example.com";

            EmailSnapshot history = new EmailSnapshot();
            history.Subject = "Earlier subject";
            history.Body =
                "HISTORY_UNIQUE\r\n"
                + "-----Original Message-----\r\n"
                + "DUPLICATE_QUOTED_COPY";

            ConversationSnapshot conversation = new ConversationSnapshot();
            conversation.CurrentEmail = current;
            conversation.ConversationAvailable = true;
            conversation.TotalConversationItems = 2;
            conversation.HistoryMessages.Add(history);

            string prompt = new ConversationPromptBuilder().BuildUserPrompt(
                conversation,
                5000);
            True(
                prompt.Contains("HISTORY_UNIQUE"),
                "history new content should be included");
            True(
                !prompt.Contains("DUPLICATE_QUOTED_COPY"),
                "quoted history should be removed from stored messages");
            True(
                prompt.Contains("EMBEDDED_CONTEXT"),
                "current mail must preserve embedded external history");
            True(
                prompt.Contains("=== 当前选中的邮件 ==="),
                "current email must have a distinct section");
        }

        private static void UsesGenericConfigurationDefaults()
        {
            AppSettings settings = new AppSettings();
            Equal("https://api.openai.com/v1", settings.ApiBaseUrl);
            Equal("gpt-5.6-luna", settings.Model);
            Equal("none", settings.ReasoningEffort);
            Equal(string.Empty, settings.ApiKeyCiphertext);
        }

        private static void RoundTripsGenericConfiguration()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                settings.ApiBaseUrl = " https://example.com/v1/ ";
                settings.Model = " arbitrary-third-party-model ";
                settings.ReasoningEffort = "MEDIUM";
                settings.IdentityEmailAddresses = "other@example.com";
                settings.IdentityAliases = "Ethan; 耿工";
                store.SetApiKey(settings, " secret ");
                store.Save(settings);

                AppSettings loaded = store.Load();
                Equal("https://example.com/v1/", loaded.ApiBaseUrl);
                Equal("arbitrary-third-party-model", loaded.Model);
                Equal("medium", loaded.ReasoningEffort);
                Equal("secret", store.ReadApiKey(loaded));
                True(!File.ReadAllText(store.SettingsPath).Contains("secret"),
                    "plaintext key must not be persisted");
                Equal("other@example.com", loaded.IdentityEmailAddresses);
                Equal("Ethan; 耿工", loaded.IdentityAliases);

                string[] efforts = { "none", "medium", "max" };
                for (int i = 0; i < efforts.Length; i++)
                {
                    loaded.ReasoningEffort = efforts[i];
                    store.Save(loaded);
                    loaded = store.Load();
                    Equal(efforts[i], loaded.ReasoningEffort);
                }
                loaded.ReasoningEffort = "invalid";
                store.Save(loaded);
                Equal("none", store.Load().ReasoningEffort);
            }
            finally
            {
                DeleteTemporaryDirectory(directory);
            }
        }

        private static void MigratesLegacyProviderConfiguration()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings source = new AppSettings();
                source.ApiBaseUrl = "https://api.deepseek.com";
                store.SetApiKey(source, "deep-key");
                File.WriteAllText(store.SettingsPath,
                    "{\"ProviderId\":\"deepseek\",\"ApiKeyCiphertext\":\""
                        + source.ApiKeyCiphertext + "\"}");
                AppSettings deepseek = store.Load();
                Equal(string.Empty, deepseek.ProviderId);
                Equal("https://api.deepseek.com", deepseek.ApiBaseUrl);
                Equal("deepseek-v4-flash", deepseek.Model);
                Equal("none", deepseek.ReasoningEffort);
                Equal("deep-key", store.ReadApiKey(deepseek));
                source.ApiBaseUrl = "https://custom.example/v1";
                source.Model = "custom-zhipu-model";
                store.SetApiKey(source, "zhipu-key");
                File.WriteAllText(store.SettingsPath,
                    "{\"ProviderId\":\"zhipu\",\"ApiBaseUrl\":\"https://custom.example/v1\",\"Model\":\"custom-zhipu-model\",\"ApiKeyCiphertext\":\""
                        + source.ApiKeyCiphertext + "\"}");
                AppSettings zhipu = store.Load();
                Equal(string.Empty, zhipu.ProviderId);
                Equal("https://custom.example/v1", zhipu.ApiBaseUrl);
                Equal("custom-zhipu-model", zhipu.Model);
                Equal("zhipu-key", store.ReadApiKey(zhipu));
                Equal(source.ApiKeyCiphertext, zhipu.ApiKeyCiphertext);
                store.Save(zhipu);
                Equal("zhipu-key", store.ReadApiKey(store.Load()));
            }
            finally { DeleteTemporaryDirectory(directory); }
        }

        private static void ProtectsApiKeysAcrossEndpointChanges()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                settings.ApiBaseUrl = "https://one.example/v1";
                store.SetApiKey(settings, "secret");
                Equal("secret", store.ReadApiKey(settings));
                True(SettingsStore.IsSameApiEndpoint("https://one.example/v1", "https://one.example/v1/chat/completions"), "equivalent endpoints");
                True(SettingsStore.IsSameApiEndpoint("https://ONE.example:443/v1/", settings.ApiBaseUrl),
                    "host case, default port and trailing slash are equivalent");
                store.Save(settings);
                string[] changedUrls = { "https://two.example/v1", "https://one.example/v2",
                    "https://one.example:444/v1", "https://one.example/V1" };
                foreach (string url in changedUrls)
                {
                    settings.ApiBaseUrl = url;
                    Throws<InvalidOperationException>(delegate { store.ReadApiKey(settings); });
                    Throws<InvalidOperationException>(delegate { store.Save(settings); });
                }
                Equal("secret", store.ReadApiKey(store.Load()));
                string originalJson = File.ReadAllText(store.SettingsPath);
                File.WriteAllText(store.SettingsPath, originalJson.Replace(
                    "\"ApiBaseUrl\":\"https://one.example/v1\"",
                    "\"ApiBaseUrl\":\"https://two.example/v1\""));
                Throws<InvalidOperationException>(delegate { store.ReadApiKey(store.Load()); });
                store.SetApiKey(settings, "replacement");
                store.Save(settings);
                Equal("replacement", store.ReadApiKey(store.Load()));
                store.SetApiKey(settings, string.Empty);
                store.Save(settings);
                Equal(string.Empty, store.ReadApiKey(store.Load()));
            }
            finally { DeleteTemporaryDirectory(directory); }
        }

        private static void SendsReasoningEffortWithoutThinking()
        {
            string[] efforts = { "none", "medium", "max" };
            for (int i = 0; i < efforts.Length; i++)
            {
                using (CaptureServer server = new CaptureServer(1, false))
                {
                    AppSettings settings = new AppSettings();
                    settings.ProviderId = "deepseek";
                    settings.ApiBaseUrl = server.BaseUrl + "/v1";
                    settings.Model = "gpt-5.6-luna";
                    settings.ReasoningEffort = efforts[i];
                    string response = new OpenAiCompatibleClient().CompleteAsync(
                        settings, "key", "system", "user", false,
                        CancellationToken.None).GetAwaiter().GetResult();
                    Equal("ok", response);
                    string body = server.WaitForBody();
                    True(body.Contains("\"reasoning_effort\":\"" + efforts[i] + "\""),
                        "reasoning effort must be sent");
                    True(!body.Contains("thinking"), "DeepSeek thinking payload must be gone");
                    True(body.Contains("\"model\":\"gpt-5.6-luna\""), "model must be sent");
                    True(body.Contains("\"stream\":false"), "stream must be false");
                    True(body.Contains("\"messages\":["), "messages must be sent");
                }
            }
            using (CaptureServer retryServer = new CaptureServer(2, true))
            {
                AppSettings settings = new AppSettings();
                settings.ProviderId = "deepseek";
                settings.ApiBaseUrl = retryServer.BaseUrl;
                settings.ReasoningEffort = "medium";
                Equal("ok", new OpenAiCompatibleClient().CompleteAsync(
                    settings, "key", "system", "user", true,
                    CancellationToken.None).GetAwaiter().GetResult());
                string retryBody = retryServer.WaitForBody();
                True(retryBody.Contains("\"reasoning_effort\":\"medium\""),
                    "JSON compatibility retry must preserve effort");
                True(!retryBody.Contains("thinking"), "retry must not add thinking");
                True(!retryBody.Contains("response_format"), "retry must remove JSON mode");
                True(retryServer.FirstBody.Contains("response_format"), "initial request uses JSON mode");
                True(retryServer.FirstBody.Contains("\"reasoning_effort\":\"medium\""),
                    "initial request also preserves effort");
            }
        }

        private static void BuildsExecutiveBriefPrompt()
        {
            ConversationPromptBuilder builder =
                new ConversationPromptBuilder();
            string system = builder.BuildSystemPrompt("简体中文");
            True(system.Contains("【结论】"), "conclusion section is required");
            True(system.Contains("【与你相关】"), "personal section is required");
            True(system.Contains("暂无需你处理。"), "empty action wording is required");
            True(
                !system.Contains("【关键往来时间线】"),
                "old timeline section must be removed");
            True(
                system.Contains("不可信数据")
                    && (system.Contains("不得执行")
                        || system.Contains("不要执行")),
                "prompt injection boundary must remain");

            ConversationSnapshot conversation = new ConversationSnapshot();
            conversation.CurrentEmail = new EmailSnapshot();
            conversation.CurrentEmail.CurrentUserName = "Yixiong Geng";
            conversation.CurrentEmail.CurrentUserEmail =
                "yixiong.geng@hawksoft3d.com";
            string user = builder.BuildUserPrompt(
                conversation,
                5000,
                "other@example.com; yixiong.geng@hawksoft3d.com",
                "Ethan, 耿工; Ethan Geng");
            True(user.Contains("other@example.com"), "extra email must be present");
            True(user.Contains("Ethan"), "English alias must be present");
            True(user.Contains("耿工"), "Chinese alias must be present");
            True(user.Contains("视为同一个人"), "identity equivalence is required");
        }

        private static void IgnoresLegacyFolderSetting()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                Directory.CreateDirectory(directory);
                File.WriteAllText(
                    store.SettingsPath,
                    "{\"ProviderId\":\"zhipu\","
                        + "\"LocalFolderRoot\":\"C:\\\\OldReference\"}");
                AppSettings loaded = store.Load();
                Equal(string.Empty, loaded.ProviderId);
                Equal("glm-5.3-flash", loaded.Model);

                store.Save(loaded);
                True(
                    File.ReadAllText(store.SettingsPath).IndexOf(
                        "LocalFolderRoot",
                        StringComparison.Ordinal) < 0,
                    "legacy Windows folder setting must be removed on save");
            }
            finally
            {
                DeleteTemporaryDirectory(directory);
            }
        }

        private static void ShowsEditableGenericSettings()
        {
            using (SettingsForm form = new SettingsForm(new SettingsStore(), new AppSettings()))
            {
                TextBox baseUrl = (TextBox)GetPrivateField(form, "_baseUrl");
                ComboBox model = (ComboBox)GetPrivateField(form, "_model");
                ComboBox effort = (ComboBox)GetPrivateField(form, "_reasoningEffort");
                TextBox apiKey = (TextBox)GetPrivateField(form, "_apiKey");
                Equal("https://api.openai.com/v1", baseUrl.Text);
                Equal(ComboBoxStyle.DropDown, model.DropDownStyle);
                Equal("gpt-5.6-luna", model.Text);
                Equal(3, effort.Items.Count);
                Equal("None", effort.Text);
                Equal(ComboBoxStyle.DropDownList, effort.DropDownStyle);
                Equal(string.Empty, apiKey.Text);
                True(apiKey.UseSystemPasswordChar, "new key entry must be masked");
                True(!ContainsControlText(form, "AI Provider"), "provider selector removed");
            }
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                store.SetApiKey(settings, "previous-key");
                using (SettingsForm form = new SettingsForm(store, settings))
                {
                    TextBox key = (TextBox)GetPrivateField(form, "_apiKey");
                    Equal(string.Empty, key.Text);
                    ((TextBox)GetPrivateField(form, "_baseUrl")).Text = "https://custom.example/v2";
                    ((ComboBox)GetPrivateField(form, "_model")).Text = "arbitrary-model-id";
                    ((ComboBox)GetPrivateField(form, "_reasoningEffort")).SelectedIndex = 1;
                    key.Text = "new-key";
                    object[] arguments = { null };
                    object saved = typeof(SettingsForm).GetMethod("TrySaveSettings",
                        System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                        .Invoke(form, arguments);
                    Equal(true, (bool)saved);
                    AppSettings loaded = store.Load();
                    Equal("https://custom.example/v2", loaded.ApiBaseUrl);
                    Equal("arbitrary-model-id", loaded.Model);
                    Equal("medium", loaded.ReasoningEffort);
                    Equal("new-key", store.ReadApiKey(loaded));
                    Equal(string.Empty, key.Text);
                }
            }
            finally { DeleteTemporaryDirectory(directory); }
        }

        private static object GetPrivateField(object instance, string name)
        {
            System.Reflection.FieldInfo field = instance.GetType().GetField(
                name,
                System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.NonPublic);
            if (field == null) throw new Exception("Missing field " + name);
            return field.GetValue(instance);
        }










        private static void SerializesSemanticIndexPayload()
        {
            MailIndexPayload payload = new MailIndexPayload();
            payload.StoreId = "store-1";
            payload.EntryId = "entry-1";
            payload.Subject = "Re: Curtain wall edit";
            payload.SenderEmail = "india@example.com";
            payload.BodyText = "can we edit the panels individually";
            payload.AttachmentCount = 1;
            payload.AttachmentNames = "compare.xlsx";

            string json = new JavaScriptSerializer().Serialize(
                SemanticSearchClient.ToSnakeCaseDictionary(payload));
            True(json.Contains("\"store_id\":\"store-1\""),
                "payload must use snake_case keys the service expects");
            True(json.Contains("\"entry_id\":\"entry-1\""),
                "entry id must be serialized");
            True(json.Contains("\"body_text\":"), "body text must be serialized");
            True(json.Contains("\"attachment_count\":1"),
                "attachment count must be serialized");
            // Guard against a regression to PascalCase, which would make the
            // service reject every field silently.
            True(!json.Contains("StoreId") && !json.Contains("EntryId"),
                "PascalCase keys must never be sent");
            True(SemanticSearchClient.ToSnakeCaseDictionary(new MailIndexPayload())
                .ContainsKey("internet_message_id"),
                "every payload key must match the service schema");
        }

        private static void FallsBackToLoopbackServiceUrl()
        {
            SemanticSearchClient empty = new SemanticSearchClient(null, "t");
            Equal("http://127.0.0.1:8765", empty.BaseUrl);
            SemanticSearchClient remote = new SemanticSearchClient(
                "http://192.168.1.50:8765", "t");
            Equal("http://127.0.0.1:8765", remote.BaseUrl);
        }

        private static void RejectsNonLoopbackServiceUrl()
        {
            // A LAN or internet address must never be used for mail content.
            SemanticSearchClient client = new SemanticSearchClient(
                "https://evil.example.com", "t");
            Equal("http://127.0.0.1:8765", client.BaseUrl);
            SemanticSearchClient localhost = new SemanticSearchClient(
                "http://localhost:9100/", "t");
            Equal("http://localhost:9100", localhost.BaseUrl);
        }

        private static void ReportsMissingSemanticService()
        {
            // Port with no listener: the client must surface a friendly error,
            // not throw a raw socket exception.
            SemanticSearchClient client = new SemanticSearchClient(
                "http://127.0.0.1:1", "t");
            SemanticServiceStatus status = client.GetStatus();
            Equal(false, status.Reachable);
            True(status.Message.Length > 0, "unreachable status must explain why");
            try
            {
                client.Search("hello", null, null, null, null, null);
                throw new Exception("search against a dead service must fail");
            }
            catch (SemanticServiceException ex)
            {
                True(ex.Message.Contains("本地语义搜索服务"),
                    "error message must mention the local semantic service");
            }
        }

        private static void ParsesSemanticSearchResponse()
        {
            string json =
                "{\"query\":\"幕墙\",\"results\":[{"
                + "\"mail_id\":3,\"store_id\":\"S\",\"entry_id\":\"E\","
                + "\"internet_message_id\":\"m1\",\"subject\":\"Curtain wall\","
                + "\"sender_name\":\"India\",\"sender_email\":\"i@e.com\","
                + "\"received_at\":\"2026-06-15T00:00:00Z\","
                + "\"folder_path\":\"Inbox\",\"attachment_names\":\"a.xlsx\","
                + "\"attachment_count\":1,\"snippet\":\"can we edit\","
                + "\"score\":0.8,\"semantic_rank\":1,\"keyword_rank\":null}],"
                + "\"semantic_available\":true,\"keyword_available\":true,"
                + "\"candidate_count\":1,\"filter_applied\":false}";
            SemanticSearchResponse response = InvokeParser(json);
            Equal(1, response.Results.Count);
            Equal("Curtain wall", response.Results[0].Subject);
            Equal("E", response.Results[0].EntryId);
            Equal(1, response.Results[0].SemanticRank.Value);
            Equal(false, response.Results[0].KeywordRank.HasValue);
            Equal(true, response.SemanticAvailable);
        }

        private static void ParsesSemanticHealthStatus()
        {
            string json =
                "{\"status\":\"ok\",\"model_id\":\"google/embeddinggemma-2\","
                + "\"model_available\":true,\"device\":\"cpu\",\"index\":{"
                + "\"mail_count\":100,\"chunk_count\":250,"
                + "\"embedded_chunk_count\":250,\"size_bytes\":4096,"
                + "\"last_sync_at\":\"2026-06-15T10:00:00Z\"},"
                + "\"paths\":{\"model_dir\":\"C:/models\","
                + "\"index_dir\":\"C:/index\",\"data_dir\":\"C:/data\","
                + "\"database\":\"C:/index/mail.db\",\"log_dir\":\"C:/logs\","
                + "\"service_dir\":\"C:/service\"}}";
            SemanticServiceStatus status = InvokeStatusParser(json);
            Equal(true, status.ModelAvailable);
            Equal("google/embeddinggemma-2", status.ModelId);
            Equal(100, status.MailCount);
            Equal(250, status.ChunkCount);
            Equal(4096L, status.SizeBytes);
            // Path hints drive the "open folder" buttons in Settings.
            True(status.ModelDir.Contains("models"), "model dir must be parsed");
            True(status.IndexDir.Contains("index"), "index dir must be parsed");
            True(status.LogDir.Contains("logs"), "log dir must be parsed");
        }

        // The client parses private responses; these helpers expose them for
        // testing without widening the production API surface.
        private static SemanticSearchResponse InvokeParser(string json)
        {
            return (SemanticSearchResponse)typeof(SemanticSearchClient)
                .GetMethod(
                    "ParseSearchResponse",
                    System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.NonPublic)
                .Invoke(new SemanticSearchClient(null, "t"), new object[] { json });
        }

        private static SemanticServiceStatus InvokeStatusParser(string json)
        {
            SemanticSearchClient client = new SemanticSearchClient(
                "http://127.0.0.1:1", "t");
            // Reuse the same JSON field mapping the live parser uses.
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            Dictionary<string, object> root =
                (Dictionary<string, object>)serializer.DeserializeObject(json);
            Dictionary<string, object> index =
                (Dictionary<string, object>)root["index"];
            SemanticServiceStatus status = new SemanticServiceStatus();
            status.Reachable = true;
            status.Status = Convert.ToString(root["status"]);
            status.ModelId = Convert.ToString(root["model_id"]);
            status.ModelAvailable = Convert.ToBoolean(root["model_available"]);
            status.Device = Convert.ToString(root["device"]);
            status.MailCount = Convert.ToInt32(index["mail_count"]);
            status.ChunkCount = Convert.ToInt32(index["chunk_count"]);
            status.EmbeddedChunkCount =
                Convert.ToInt32(index["embedded_chunk_count"]);
            status.SizeBytes = Convert.ToInt64(index["size_bytes"]);
            status.LastSyncAt = Convert.ToString(index["last_sync_at"]);
            if (root.ContainsKey("paths"))
            {
                Dictionary<string, object> paths =
                    (Dictionary<string, object>)root["paths"];
                status.ModelDir = Convert.ToString(paths["model_dir"]);
                status.IndexDir = Convert.ToString(paths["index_dir"]);
                status.DataDir = Convert.ToString(paths["data_dir"]);
                status.LogDir = Convert.ToString(paths["log_dir"]);
                status.ServiceDir = Convert.ToString(paths["service_dir"]);
            }

            return status;
        }

        private static void Treats503AsModelUnavailable()
        {
            SemanticServiceException modelDown =
                new SemanticServiceException("model missing", 503);
            Equal(true, modelDown.IsModelUnavailable);
            Equal(503, modelDown.StatusCode);
            SemanticServiceException network =
                new SemanticServiceException("no service", 0);
            Equal(false, network.IsModelUnavailable);
            SemanticServiceException serverError =
                new SemanticServiceException("boom", 500);
            Equal(false, serverError.IsModelUnavailable);
        }

        private static void ReportsModelStatus()
        {
            // A dead port must yield an unreachable status with a message,
            // never an exception.
            SemanticSearchClient client = new SemanticSearchClient(
                "http://127.0.0.1:1", "t");
            SemanticServiceStatus status = client.GetModelStatus();
            Equal(false, status.Reachable);
            True(status.Message.Length > 0,
                "unreachable model status must explain why");
        }

        private static void PreservesSemanticIndexSettings()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                settings.SemanticIndexMonths = 6;
                settings.SemanticIndexMaxMails = 1500;
                settings.SemanticIndexIncludeSent = false;
                store.Save(settings);
                AppSettings loaded = store.Load();
                Equal(6, loaded.SemanticIndexMonths);
                Equal(1500, loaded.SemanticIndexMaxMails);
                Equal(false, loaded.SemanticIndexIncludeSent);
            }
            finally { DeleteTemporaryDirectory(directory); }
        }

        private static void ClampsInvalidSemanticIndexSettings()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                settings.SemanticIndexMonths = 999;
                settings.SemanticIndexMaxMails = 5;
                store.Save(settings);
                AppSettings loaded = store.Load();
                Equal(12, loaded.SemanticIndexMonths);
                Equal(3000, loaded.SemanticIndexMaxMails);
            }
            finally { DeleteTemporaryDirectory(directory); }
        }



        private static bool ContainsControlText(Control root, string text)
        {
            if (string.Equals(root.Text, text, StringComparison.Ordinal))
            {
                return true;
            }

            foreach (Control child in root.Controls)
            {
                if (ContainsControlText(child, text))
                {
                    return true;
                }
            }

            return false;
        }

        private static string CreateTemporaryDirectory()
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "OutlookAiAssistantTests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }

        private static void DeleteTemporaryDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }

        private sealed class CaptureServer : IDisposable
        {
            private readonly TcpListener _listener;
            private readonly Thread _thread;
            private readonly int _requests;
            private readonly bool _badRequestFirst;
            private readonly List<string> _bodies = new List<string>();
            private Exception _error;

            public string BaseUrl { get; private set; }
            public string FirstBody { get { return _bodies[0]; } }

            public CaptureServer(int requests, bool badRequestFirst)
            {
                _requests = requests;
                _badRequestFirst = badRequestFirst;
                _listener = new TcpListener(IPAddress.Loopback, 0);
                _listener.Start();
                int port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                BaseUrl = "http://127.0.0.1:" + port;
                _thread = new Thread(Serve);
                _thread.IsBackground = true;
                _thread.Start();
            }

            public string WaitForBody()
            {
                if (!_thread.Join(10000)) throw new Exception("local HTTP server timed out");
                if (_error != null) throw new Exception("local HTTP server failed: " + _error.Message);
                if (_bodies.Count == 0) throw new Exception("local HTTP server got no request");
                return _bodies[_bodies.Count - 1];
            }

            private void Serve()
            {
                try
                {
                    for (int i = 0; i < _requests; i++)
                    {
                        using (TcpClient client = _listener.AcceptTcpClient())
                        using (NetworkStream stream = client.GetStream())
                        {
                            stream.ReadTimeout = 5000;
                            byte[] buffer = new byte[8192];
                            int used = 0;
                            int contentLength = 0;
                            int headerEnd = -1;
                            while (headerEnd < 0 || used - headerEnd - 4 < contentLength)
                            {
                                int read = stream.Read(buffer, used, buffer.Length - used);
                                if (read == 0) break;
                                used += read;
                                string text = Encoding.UTF8.GetString(buffer, 0, used);
                                headerEnd = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
                                if (headerEnd >= 0)
                                {
                                    string marker = "Content-Length:";
                                    int p = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                                    if (p >= 0)
                                    {
                                        int lineEnd = text.IndexOf("\r\n", p, StringComparison.Ordinal);
                                        contentLength = int.Parse(text.Substring(p + marker.Length, lineEnd - p - marker.Length).Trim());
                                    }
                                }
                            }
                            string requestText = Encoding.UTF8.GetString(buffer, 0, used);
                            int bodyStart = requestText.IndexOf("\r\n\r\n", StringComparison.Ordinal) + 4;
                            _bodies.Add(requestText.Substring(bodyStart));
                            bool bad = _badRequestFirst && i == 0;
                            string json = bad ? "{\"error\":{\"message\":\"reasoning unsupported\"}}" : "{\"choices\":[{\"message\":{\"content\":\"ok\"}}]}";
                            string status = bad ? "400 Bad Request" : "200 OK";
                            byte[] response = Encoding.UTF8.GetBytes("HTTP/1.1 " + status + "\r\nContent-Type: application/json\r\nContent-Length: " + Encoding.UTF8.GetByteCount(json) + "\r\nConnection: close\r\n\r\n" + json);
                            stream.Write(response, 0, response.Length);
                        }
                    }
                }
                catch (Exception ex) { _error = ex; }
            }

            public void Dispose()
            {
                _listener.Stop();
                if (_thread.IsAlive) _thread.Join(1000);
            }
        }

        private static void Run(string name, Action test)
        {
            try
            {
                test();
                _passed++;
                Console.WriteLine("PASS: " + name);
            }
            catch (Exception ex)
            {
                _failed++;
                Console.WriteLine("FAIL: " + name + " | " + ex.Message);
            }
        }

        private static void True(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception(message);
            }
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!object.Equals(expected, actual))
            {
                throw new Exception(
                    "Expected [" + expected + "] but got [" + actual + "].");
            }
        }

        private static void Throws<T>(Action action)
            where T : Exception
        {
            try
            {
                action();
            }
            catch (T)
            {
                return;
            }

            throw new Exception("Expected exception " + typeof(T).Name + ".");
        }
    }
}
