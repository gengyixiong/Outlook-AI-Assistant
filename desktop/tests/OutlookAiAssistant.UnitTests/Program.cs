using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.EntityIndex;
using OutlookAiAssistant.Models;
using OutlookAiAssistant.Search;
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
            Run("Parses snake_case search plan", ParsesSnakeCaseSearchPlan);
            Run("Compiles layered allow-listed AQS", CompilesLayeredAqs);
            Run("Broad search uses only the first concept without anchors", BroadUsesFirstConcept);
            Run("Broad search falls back to explicit fields", BroadFallsBackToExplicitFields);
            Run("Parses legacy text groups as concepts", ParsesLegacyTextGroups);
            Run("Demotes relationship descriptions from sender", DemotesRelationshipSender);
            Run("Sanitizes generated values", SanitizesGeneratedValues);
            Run("Rejects empty plans", RejectsEmptyPlans);
            Run("Rejects invalid date ranges", RejectsInvalidDateRanges);
            Run("Strips markdown JSON fences", StripsMarkdownJsonFences);
            Run("Builds compatible endpoint URLs", BuildsCompatibleEndpointUrls);
            Run("Builds settings dialog before first display", BuildsSettingsDialog);
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
            Run("Round-trips contact index JSON", RoundTripsContactIndex);
            Run("Backs up contact index safely", BacksUpContactIndex);
            Run("Rejects malformed contact index", RejectsMalformedContactIndex);
            Run("Creates stable unique contact IDs", CreatesStableContactIds);
            Run("Stores Outlook folder contact labels", StoresOutlookFolderLabels);
            Run("Resolves validated contact IDs", ResolvesValidatedContactIds);
            Run("Drops unknown contact IDs", DropsUnknownContactIds);
            Run("Compiles attachment search", CompilesAttachmentSearch);
            Run("Folder context cannot become sender", FolderContextCannotBecomeSender);

            Console.WriteLine(
                "Tests complete. Passed: {0}; Failed: {1}",
                _passed,
                _failed);
            return _failed == 0 ? 0 : 1;
        }

        private static void ParsesSnakeCaseSearchPlan()
        {
            const string json =
                "{"
                + "\"from\":[\"Richard\"],"
                + "\"to\":[],\"cc\":[],\"to_me\":true,"
                + "\"anchor_groups\":[[\"GstarBIM\",\"Gstar BIM\"]],"
                + "\"concept_groups\":[[\"大赛\",\"competition\"]],"
                + "\"hint_groups\":[[\"Brazil\",\"Brasil\"]],"
                + "\"subject_groups\":[],\"body_groups\":[],"
                + "\"received_from\":\"2026-06-01\","
                + "\"received_through\":\"2026-06-30\","
                + "\"has_attachments\":true,\"is_unread\":null,"
                + "\"scope\":\"all_folders\""
                + "}";

            SearchPlan plan = new SearchPlanParser().Parse(json);
            Equal("Richard", plan.From[0]);
            True(plan.ToMe, "to_me should be true");
            Equal(1, plan.AnchorGroups.Count);
            Equal(2, plan.AnchorGroups[0].Count);
            Equal(1, plan.ConceptGroups.Count);
            Equal(1, plan.HintGroups.Count);
            Equal(true, plan.HasAttachments.Value);
            Equal("2026-06-30", plan.ReceivedThrough);
        }

        private static void CompilesLayeredAqs()
        {
            SearchPlan plan = new SearchPlan();
            plan.From.Add("Richard");
            plan.AnchorGroups.Add(new System.Collections.Generic.List<string>
            {
                "GstarBIM",
                "Gstar BIM"
            });
            plan.AnchorGroups.Add(new System.Collections.Generic.List<string>
            {
                "second anchor"
            });
            plan.ConceptGroups.Add(new System.Collections.Generic.List<string>
            {
                "competition",
                "contest"
            });
            plan.HintGroups.Add(new System.Collections.Generic.List<string>
            {
                "Brazil",
                "Brasil"
            });
            plan.ReceivedFrom = "2026-06-01";
            plan.ReceivedThrough = "2026-06-30";
            plan.HasAttachments = true;

            SearchQuerySet queries = new AqsQueryCompiler().CompileAll(plan);
            Equal(
                "(\"GstarBIM\" OR \"Gstar BIM\")",
                queries.Broad);
            Equal(
                "(\"GstarBIM\" OR \"Gstar BIM\")"
                + " AND \"second anchor\""
                + " AND (\"competition\" OR \"contest\")"
                + " AND from:\"Richard\""
                + " AND received:>=2026-06-01"
                + " AND received:<=2026-06-30"
                + " AND hasattachments:yes",
                queries.Recommended);
            Equal(
                "(\"GstarBIM\" OR \"Gstar BIM\")"
                + " AND \"second anchor\""
                + " AND (\"competition\" OR \"contest\")"
                + " AND (\"Brazil\" OR \"Brasil\")"
                + " AND from:\"Richard\""
                + " AND received:>=2026-06-01"
                + " AND received:<=2026-06-30"
                + " AND hasattachments:yes",
                queries.Precise);
        }

        private static void BroadUsesFirstConcept()
        {
            SearchPlan plan = new SearchPlan();
            plan.ConceptGroups.Add(
                new System.Collections.Generic.List<string>
                {
                    "competition",
                    "contest"
                });
            plan.ConceptGroups.Add(
                new System.Collections.Generic.List<string>
                {
                    "Brazil",
                    "Brasil"
                });

            string query = new AqsQueryCompiler().Compile(plan);
            Equal("(\"competition\" OR \"contest\")", query);
        }

        private static void BroadFallsBackToExplicitFields()
        {
            SearchPlan plan = new SearchPlan();
            plan.From.Add("Richard");
            plan.HasAttachments = true;

            string query = new AqsQueryCompiler().Compile(plan);
            Equal("from:\"Richard\" AND hasattachments:yes", query);
        }

        private static void ParsesLegacyTextGroups()
        {
            SearchPlan plan = new SearchPlanParser().Parse(
                "{\"text_groups\":[[\"quotation\",\"quote\"]]}");
            Equal(1, plan.ConceptGroups.Count);
            Equal("quotation", plan.ConceptGroups[0][0]);
        }

        private static void DemotesRelationshipSender()
        {
            SearchPlan plan = new SearchPlanParser().Parse(
                "{\"from\":[\"巴西客户\",\"Richard Li\"],"
                + "\"hint_groups\":[]}");
            Equal(1, plan.From.Count);
            Equal("Richard Li", plan.From[0]);
            Equal(1, plan.HintGroups.Count);
            Equal("巴西客户", plan.HintGroups[0][0]);
        }

        private static void SanitizesGeneratedValues()
        {
            SearchPlan plan = new SearchPlan();
            plan.SubjectGroups.Add(
                new System.Collections.Generic.List<string>
                {
                    "quote\") OR from:(\"boss"
                });

            string query = new AqsQueryCompiler().Compile(
                plan,
                SearchStrictness.Recommended);
            True(!query.Contains("\")"), "quotes and parentheses must be removed");
            True(
                query.StartsWith("subject:\"", StringComparison.Ordinal),
                "term must stay inside a field value");
        }

        private static void RejectsEmptyPlans()
        {
            Throws<InvalidOperationException>(
                delegate { new AqsQueryCompiler().Compile(new SearchPlan()); });
        }

        private static void RejectsInvalidDateRanges()
        {
            const string json =
                "{\"received_from\":\"2026-07-20\","
                + "\"received_through\":\"2026-07-01\"}";
            Throws<InvalidOperationException>(
                delegate { new SearchPlanParser().Parse(json); });
        }

        private static void StripsMarkdownJsonFences()
        {
            SearchPlan plan = new SearchPlanParser().Parse(
                "```json\n{\"from\":[\"Laszlo\"],\"scope\":\"current_folder\"}\n```");
            Equal("Laszlo", plan.From[0]);
            Equal("current_folder", plan.Scope);
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

        private static void RoundTripsContactIndex()
        {
            ContactIndexStore store = new ContactIndexStore(
                CreateTemporaryDirectory());
            try
            {
                ContactIndex index = CreateContactIndex(
                    "jieteng.luo@hawksoft3d.com",
                    "Jason Luo");
                index.Contacts[0].Aliases.Add("Jason");
                FolderContextEntry folder = new FolderContextEntry
                {
                    Root = "Mailbox - Yixiong",
                    RelativePath = "13. 土耳其代理商 system24"
                };
                folder.ContactIds.Add(index.Contacts[0].Id);
                folder.ContactIds.Add("contact_does_not_exist");
                index.FolderContext.Add(folder);
                string json = store.Serialize(index);
                True(json.Contains("\"created_at\""), "created_at must be snake_case");
                True(json.Contains("\"folder_context\""), "folder_context must share the JSON");
                True(json.Contains("\"contact_ids\""), "folder labels need contact IDs");

                ContactIndex parsed = store.Parse(json);
                Equal(1, parsed.Contacts.Count);
                Equal("jieteng.luo@hawksoft3d.com", parsed.Contacts[0].Email);
                True(parsed.Contacts[0].Aliases.Contains("Jason"), "alias must survive");
                True(
                    parsed.Contacts[0].Aliases.Contains("jieteng.luo"),
                    "email local part must be an alias");
                Equal(1, parsed.FolderContext.Count);
                Equal(
                    index.Contacts[0].Id,
                    parsed.FolderContext[0].ContactIds[0]);
                Equal(1, parsed.FolderContext[0].ContactIds.Count);
            }
            finally
            {
                DeleteTemporaryDirectory(store.DirectoryPath);
            }
        }

        private static void BacksUpContactIndex()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                ContactIndexStore store = new ContactIndexStore(directory);
                store.Save(CreateContactIndex(
                    "kun.ma@hawksoft3d.com",
                    "Kun Ma"));
                string first = File.ReadAllText(store.IndexPath);
                store.Save(CreateContactIndex(
                    "jieteng.luo@hawksoft3d.com",
                    "Jason Luo"));
                Equal(1, Directory.GetFiles(
                    store.BackupDirectory,
                    "contact-index-*.json").Length);

                ContactIndex invalid = CreateContactIndex(
                    "first@example.com",
                    "First");
                ContactProfile duplicate = new ContactProfile();
                duplicate.Id = invalid.Contacts[0].Id;
                duplicate.Email = "second@example.com";
                invalid.Contacts.Add(duplicate);
                string beforeFailure = File.ReadAllText(store.IndexPath);
                Throws<InvalidOperationException>(delegate { store.Save(invalid); });
                Equal(beforeFailure, File.ReadAllText(store.IndexPath));
                True(first.Length > 0, "first index should have been written");
            }
            finally
            {
                DeleteTemporaryDirectory(directory);
            }
        }

        private static void StoresOutlookFolderLabels()
        {
            ContactIndexStore store = new ContactIndexStore(
                CreateTemporaryDirectory());
            try
            {
                ContactIndex index = CreateContactIndex(
                    "lais@example.com",
                    "Lais Stefany");
                FolderContextEntry folder = new FolderContextEntry();
                folder.Root = "Mailbox - Yixiong";
                folder.RelativePath = "Inbox\\客户\\巴西代理商 Lais Stefany";
                folder.ContactIds.Add(index.Contacts[0].Id);
                index.FolderContext.Add(folder);

                ContactIndex parsed = store.Parse(store.Serialize(index));
                Equal(1, parsed.FolderContext.Count);
                Equal(folder.RelativePath, parsed.FolderContext[0].RelativePath);
                Equal(
                    index.Contacts[0].Id,
                    parsed.FolderContext[0].ContactIds[0]);
            }
            finally
            {
                DeleteTemporaryDirectory(store.DirectoryPath);
            }
        }

        private static void RejectsMalformedContactIndex()
        {
            ContactIndexStore store = new ContactIndexStore(
                CreateTemporaryDirectory());
            try
            {
                Throws<InvalidOperationException>(delegate
                {
                    store.Parse(
                        "{\"version\":1,\"created_at\":\"2026-08-27T10:00:00\","
                        + "\"contacts\":{},\"folder_context\":[]}");
                });

                ContactIndex index = CreateContactIndex(
                    "known@example.com",
                    "Known");
                index.Contacts[0].Id = "contact_forged";
                Throws<InvalidOperationException>(delegate
                {
                    store.Serialize(index);
                });
            }
            finally
            {
                DeleteTemporaryDirectory(store.DirectoryPath);
            }
        }

        private static void CreatesStableContactIds()
        {
            string first = ContactProfile.CreateId("Kun.Ma@HawkSoft3D.com");
            string same = ContactProfile.CreateId(" kun.ma@hawksoft3d.com ");
            string other = ContactProfile.CreateId(
                "jieteng.luo@hawksoft3d.com");
            Equal(first, same);
            True(first != other, "different emails need different IDs");
        }

        private static void ResolvesValidatedContactIds()
        {
            ContactIndex index = new ContactIndex();
            index.CreatedAt = DateTime.Now.ToString("s");
            ContactProfile ma = CreateContact(
                "kun.ma@hawksoft3d.com",
                "Kun Ma",
                "马坤");
            ContactProfile jason = CreateContact(
                "jieteng.luo@hawksoft3d.com",
                "Jason Luo",
                "Jason");
            ContactProfile richard = CreateContact(
                "richard@example.com",
                "Richard",
                "Richard");
            index.Contacts.Add(ma);
            index.Contacts.Add(jason);
            index.Contacts.Add(richard);

            SearchPlan plan = new SearchPlan();
            plan.MatchedContactIds.Add(ma.Id);
            plan.MatchedContactIds.Add(jason.Id);
            plan.MatchedContactIds.Add(richard.Id);
            new ContactIndexResolver().Resolve(plan, index);
            Equal(3, plan.MatchedContactIds.Count);
            True(plan.From.Contains(ma.Email), "Ma email must resolve locally");
            True(plan.From.Contains(jason.Email), "Jason email must resolve locally");
            True(plan.From.Contains(richard.Email), "Richard email must resolve locally");
        }

        private static void DropsUnknownContactIds()
        {
            ContactIndex index = CreateContactIndex(
                "known@example.com",
                "Known");
            SearchPlan plan = new SearchPlan();
            plan.MatchedContactIds.Add("contact_does_not_exist");
            new ContactIndexResolver().Resolve(plan, index);
            Equal(0, plan.MatchedContactIds.Count);
            Equal(0, plan.From.Count);
        }

        private static void CompilesAttachmentSearch()
        {
            ContactIndex index = CreateContactIndex(
                "kun.ma@hawksoft3d.com",
                "Kun Ma");
            string contactId = index.Contacts[0].Id;
            SearchPlan plan = new SearchPlanParser().Parse(
                "{\"matched_contact_ids\":[\"" + contactId + "\"],"
                + "\"attachment_name_groups\":[[\"LT vs Pro\",\"LT PRO\"]],"
                + "\"attachment_extensions\":[\".XLSX\",\"xls\"],"
                + "\"has_attachments\":true}");
            new ContactIndexResolver().Resolve(plan, index);
            string query = new AqsQueryCompiler().Compile(
                plan,
                SearchStrictness.Recommended);
            True(
                query.Contains("from:\"kun.ma@hawksoft3d.com\""),
                "resolved email must be used");
            True(query.Contains("attachments:\"LT vs Pro\""), "filename must compile");
            True(query.Contains("attachments:\"xlsx\""), "extension must compile");
            True(query.Contains("hasattachments:yes"), "attachment filter is required");
        }

        private static void FolderContextCannotBecomeSender()
        {
            ContactIndex index = new ContactIndex();
            index.CreatedAt = DateTime.Now.ToString("s");
            index.FolderContext.Add(new FolderContextEntry
            {
                Root = "2026 H2",
                RelativePath = "13. 土耳其代理商 system24"
            });
            SearchPlan plan = new SearchPlan();
            plan.HintGroups.Add(new List<string> { "system24" });
            new ContactIndexResolver().Resolve(plan, index);
            string query = new AqsQueryCompiler().Compile(
                plan,
                SearchStrictness.Broad);
            True(
                !query.Contains("from:"),
                "folder context must never compile as a sender");
        }

        private static ContactIndex CreateContactIndex(
            string email,
            string displayName)
        {
            ContactIndex index = new ContactIndex();
            index.CreatedAt = DateTime.Now.ToString("s");
            index.Contacts.Add(CreateContact(email, displayName, displayName));
            return index;
        }

        private static ContactProfile CreateContact(
            string email,
            string displayName,
            string alias)
        {
            ContactProfile contact = new ContactProfile();
            contact.Id = ContactProfile.CreateId(email);
            contact.Email = email;
            contact.DisplayName = displayName;
            contact.Aliases.Add(alias);
            return contact;
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
