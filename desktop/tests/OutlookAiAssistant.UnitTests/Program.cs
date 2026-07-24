using System;
using OutlookAiAssistant.AI;
using OutlookAiAssistant.Configuration;
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
                "https://api.openai.com/v1/chat/completions",
                OpenAiCompatibleClient.BuildChatCompletionsUrl(
                    "https://api.openai.com/v1/chat/completions"));
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
