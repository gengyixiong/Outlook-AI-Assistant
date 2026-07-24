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
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            plan.Normalize();
            List<string> conditions = new List<string>();
            AddFieldAlternatives(conditions, "from", plan.From);
            AddFieldAlternatives(conditions, "to", plan.To);
            AddFieldAlternatives(conditions, "cc", plan.Cc);

            if (plan.ToMe)
            {
                conditions.Add("to:me");
            }

            AddTextGroups(conditions, string.Empty, plan.TextGroups);
            AddTextGroups(conditions, "subject", plan.SubjectGroups);
            AddTextGroups(conditions, "body", plan.BodyGroups);

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

            if (plan.HasAttachments.HasValue)
            {
                conditions.Add(
                    "hasattachments:"
                        + (plan.HasAttachments.Value ? "yes" : "no"));
            }

            if (plan.IsUnread.HasValue)
            {
                conditions.Add("read:" + (plan.IsUnread.Value ? "no" : "yes"));
            }

            if (conditions.Count == 0)
            {
                throw new InvalidOperationException(
                    "没有生成可执行的搜索条件，请换一种说法再试。");
            }

            return string.Join(" AND ", conditions.ToArray());
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
            IEnumerable<List<string>> groups)
        {
            foreach (List<string> group in groups.Take(12))
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

