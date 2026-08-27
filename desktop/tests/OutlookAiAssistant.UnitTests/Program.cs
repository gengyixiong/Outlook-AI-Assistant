using System;
using System.Collections.Generic;
using System.IO;
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
            Run("Builds provider URLs", BuildsProviderUrls);
            Run("Builds settings dialog before first display", BuildsSettingsDialog);
            Run("Trims quoted history from stored messages", TrimsQuotedHistory);
            Run("Keeps ordinary From lines", KeepsOrdinaryFromLines);
            Run("Builds separate history and current prompt", BuildsConversationPrompt);
            Run("Exposes only fixed Flash providers", ExposesOnlyFlashProviders);
            Run("Locks provider endpoint and model", LocksProviderConfiguration);
            Run("Ignores legacy Windows folder setting", IgnoresLegacyFolderSetting);
            Run("Builds executive brief identity prompt", BuildsExecutiveBriefPrompt);
            Run("Settings hide endpoint and model editors", SettingsHideModelEditors);
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

        private static void BuildsProviderUrls()
        {
            Equal(
                "https://api.deepseek.com/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://api.deepseek.com/"));
            Equal(
                "https://open.bigmodel.cn/api/paas/v4/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://open.bigmodel.cn/api/paas/v4/chat/completions"));
            Throws<InvalidOperationException>(
                delegate
                {
                    OpenAiCompatibleClient.BuildChatCompletionsUrl(
                        "http://example.com/v1");
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

        private static void ExposesOnlyFlashProviders()
        {
            IList<AiProviderPreset> providers =
                AiProviderPreset.CreateDefaults();
            Equal(2, providers.Count);
            Equal("deepseek", providers[0].Id);
            Equal("deepseek-v4-flash", providers[0].DefaultModel);
            Equal("zhipu", providers[1].Id);
            Equal("glm-5.3-flash", providers[1].DefaultModel);
            True(
                providers[0].DefaultModel.IndexOf(
                    "pro",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "DeepSeek preset must not expose Pro");
            True(
                providers[1].DefaultModel.IndexOf(
                    "pro",
                    StringComparison.OrdinalIgnoreCase) < 0,
                "Zhipu preset must not expose Pro");
        }

        private static void LocksProviderConfiguration()
        {
            string directory = CreateTemporaryDirectory();
            try
            {
                SettingsStore store = new SettingsStore(directory);
                AppSettings settings = new AppSettings();
                settings.ProviderId = "openai";
                settings.ApiBaseUrl = "https://example.com/v1";
                settings.Model = "expensive-pro";
                settings.ApiKeyCiphertext = "stale-key-ciphertext";
                settings.IdentityEmailAddresses = "other@example.com";
                settings.IdentityAliases = "Ethan; 耿工";
                store.Save(settings);

                AppSettings loaded = store.Load();
                Equal("deepseek", loaded.ProviderId);
                Equal("https://api.deepseek.com", loaded.ApiBaseUrl);
                Equal("deepseek-v4-flash", loaded.Model);
                Equal(string.Empty, loaded.ApiKeyCiphertext);
                Equal("other@example.com", loaded.IdentityEmailAddresses);
                Equal("Ethan; 耿工", loaded.IdentityAliases);

                loaded.ProviderId = "zhipu";
                loaded.Model = "glm-pro";
                store.Save(loaded);
                loaded = store.Load();
                Equal("https://open.bigmodel.cn/api/paas/v4", loaded.ApiBaseUrl);
                Equal("glm-5.3-flash", loaded.Model);
            }
            finally
            {
                DeleteTemporaryDirectory(directory);
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
                Equal("zhipu", loaded.ProviderId);
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

        private static void SettingsHideModelEditors()
        {
            using (SettingsForm form = new SettingsForm(
                new SettingsStore(),
                new AppSettings()))
            {
                True(
                    !ContainsControlText(form, "API 地址"),
                    "API endpoint must not be editable");
                True(
                    !ContainsControlText(form, "模型 / 接入点 ID"),
                    "model must not be editable");
                True(
                    !ContainsControlText(form, "本地参考文件夹")
                        && !ContainsControlText(form, "选择文件夹"),
                    "Windows folder picker must not remain");
            }
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
