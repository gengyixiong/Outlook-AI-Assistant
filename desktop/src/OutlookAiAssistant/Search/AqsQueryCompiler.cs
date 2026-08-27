using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// Compiles only allow-listed SearchPlan fields into Outlook AQS.
    /// Model-generated raw AQS is never executed.
    /// </summary>
    public sealed class AqsQueryCompiler
    {
        public string Compile(SearchPlan plan)
        {
            return Compile(plan, SearchStrictness.Broad);
        }

        public SearchQuerySet CompileAll(SearchPlan plan)
        {
            PreparePlan(plan);
            return new SearchQuerySet(
                CompilePrepared(plan, SearchStrictness.Broad),
                CompilePrepared(plan, SearchStrictness.Recommended),
                CompilePrepared(plan, SearchStrictness.Precise));
        }

        public string Compile(
            SearchPlan plan,
            SearchStrictness strictness)
        {
            PreparePlan(plan);
            return CompilePrepared(plan, strictness);
        }

        private static void PreparePlan(SearchPlan plan)
        {
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            plan.Normalize();
        }

        private static string CompilePrepared(
            SearchPlan plan,
            SearchStrictness strictness)
        {
            List<string> conditions = new List<string>();

            if (strictness == SearchStrictness.Broad)
            {
                AddBroadConditions(conditions, plan);
            }
            else if (strictness == SearchStrictness.Recommended
                || strictness == SearchStrictness.Precise)
            {
                AddTextGroups(
                    conditions,
                    string.Empty,
                    plan.AnchorGroups,
                    2);
                AddTextGroups(
                    conditions,
                    string.Empty,
                    GetConceptGroups(plan),
                    2);

                if (strictness == SearchStrictness.Precise)
                {
                    AddTextGroups(
                        conditions,
                        string.Empty,
                        plan.HintGroups,
                        2);
                }

                AddExplicitConditions(conditions, plan);
            }
            else
            {
                throw new ArgumentOutOfRangeException("strictness");
            }

            if (conditions.Count == 0)
            {
                throw new InvalidOperationException(
                    "没有生成可执行的搜索条件，请换一种说法再试。");
            }

            return string.Join(" AND ", conditions.ToArray());
        }

        /// <summary>
        /// The default query contains only rare anchors. If no anchor exists,
        /// it uses the first concept group without field restrictions. Explicit
        /// filters are a final fallback so a field-only plan remains usable.
        /// </summary>
        private static void AddBroadConditions(
            ICollection<string> conditions,
            SearchPlan plan)
        {
            AddTextGroups(
                conditions,
                string.Empty,
                plan.AnchorGroups,
                1);
            if (conditions.Count > 0)
            {
                return;
            }

            AddFirstTextGroup(
                conditions,
                GetConceptGroups(plan));
            if (conditions.Count == 0)
            {
                AddFirstTextGroup(conditions, plan.SubjectGroups);
            }

            if (conditions.Count == 0)
            {
                AddFirstTextGroup(conditions, plan.BodyGroups);
            }

            if (conditions.Count == 0)
            {
                AddFirstTextGroup(conditions, plan.HintGroups);
            }

            if (conditions.Count == 0)
            {
                AddExplicitConditions(conditions, plan);
            }
        }

        private static void AddExplicitConditions(
            ICollection<string> conditions,
            SearchPlan plan)
        {
            AddFieldAlternatives(conditions, "from", plan.From);
            AddFieldAlternatives(conditions, "to", plan.To);
            AddFieldAlternatives(conditions, "cc", plan.Cc);

            if (plan.ToMe)
            {
                conditions.Add("to:me");
            }

            AddTextGroups(
                conditions,
                "subject",
                plan.SubjectGroups,
                12);
            AddTextGroups(
                conditions,
                "body",
                plan.BodyGroups,
                12);
            AddTextGroups(
                conditions,
                "attachments",
                plan.AttachmentNameGroups,
                12);
            AddFieldAlternatives(
                conditions,
                "attachments",
                plan.AttachmentExtensions);

            if (!string.IsNullOrWhiteSpace(plan.ReceivedFrom))
            {
                conditions.Add(
                    "received:>=" + NormalizeDate(plan.ReceivedFrom));
            }

            if (!string.IsNullOrWhiteSpace(plan.ReceivedThrough))
            {
                conditions.Add(
                    "received:<=" + NormalizeDate(plan.ReceivedThrough));
            }

            bool hasAttachmentTerms = plan.AttachmentNameGroups.Count > 0
                || plan.AttachmentExtensions.Count > 0;
            if (hasAttachmentTerms)
            {
                conditions.Add("hasattachments:yes");
            }
            else if (plan.HasAttachments.HasValue)
            {
                conditions.Add(
                    "hasattachments:"
                        + (plan.HasAttachments.Value ? "yes" : "no"));
            }

            if (plan.IsUnread.HasValue)
            {
                conditions.Add("read:" + (plan.IsUnread.Value ? "no" : "yes"));
            }
        }

        private static void AddFieldAlternatives(
            ICollection<string> conditions,
            string field,
            IEnumerable<string> values)
        {
            List<string> safeValues = values
                .Select(NormalizeTerm)
                .Where(delegate(string value) { return value.Length > 0; })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Take(20)
                .Select(delegate(string value)
                {
                    return field + ":\"" + value + "\"";
                })
                .ToList();
            AddAlternativeGroup(conditions, safeValues);
        }

        private static void AddTextGroups(
            ICollection<string> conditions,
            string field,
            IEnumerable<List<string>> groups,
            int maximumGroups)
        {
            foreach (List<string> group in groups.Take(maximumGroups))
            {
                if (group == null)
                {
                    continue;
                }

                List<string> alternatives = group
                    .Select(NormalizeTerm)
                    .Where(delegate(string value) { return value.Length > 0; })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Take(20)
                    .Select(delegate(string value)
                    {
                        string prefix = field.Length == 0 ? string.Empty : field + ":";
                        return prefix + "\"" + value + "\"";
                    })
                    .ToList();
                AddAlternativeGroup(conditions, alternatives);
            }
        }

        private static void AddFirstTextGroup(
            ICollection<string> conditions,
            IEnumerable<List<string>> groups)
        {
            AddTextGroups(conditions, string.Empty, groups, 1);
        }

        private static IEnumerable<List<string>> GetConceptGroups(
            SearchPlan plan)
        {
            return plan.ConceptGroups.Count > 0
                ? plan.ConceptGroups
                : plan.TextGroups;
        }

        private static void AddAlternativeGroup(
            ICollection<string> conditions,
            IList<string> alternatives)
        {
            if (alternatives.Count == 1)
            {
                conditions.Add(alternatives[0]);
            }
            else if (alternatives.Count > 1)
            {
                conditions.Add("(" + string.Join(" OR ", alternatives) + ")");
            }
        }

        private static string NormalizeTerm(string value)
        {
            string safe = (value ?? string.Empty)
                .Replace("\"", " ")
                .Replace("(", " ")
                .Replace(")", " ")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            safe = Regex.Replace(safe, "\\s+", " ");
            return safe.Length <= 120 ? safe : safe.Substring(0, 120);
        }

        private static string NormalizeDate(string value)
        {
            DateTime parsed = DateTime.ParseExact(
                value,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture);
            return parsed.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }
    }
}

