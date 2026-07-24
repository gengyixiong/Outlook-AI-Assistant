using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// Parses and validates model output before any Outlook API is called.
    /// Unknown JSON fields have no effect because the compiler only reads SearchPlan.
    /// </summary>
    public sealed class SearchPlanParser
    {
        private readonly JavaScriptSerializer _serializer;

        public SearchPlanParser()
        {
            _serializer = new JavaScriptSerializer();
        }

        public SearchPlan Parse(string modelOutput)
        {
            string json = StripMarkdownFence(modelOutput);
            Dictionary<string, object> root;
            try
            {
                root = _serializer.DeserializeObject(json)
                    as Dictionary<string, object>;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "模型没有返回有效的搜索条件 JSON。",
                    ex);
            }

            if (root == null)
            {
                throw new InvalidOperationException("模型返回的搜索条件为空。");
            }

            SearchPlan plan = new SearchPlan();
            plan.From = ReadStringList(root, "from");
            plan.To = ReadStringList(root, "to");
            plan.Cc = ReadStringList(root, "cc");
            plan.ToMe = ReadBoolean(root, "to_me", false);
            plan.TextGroups = ReadGroups(root, "text_groups");
            plan.SubjectGroups = ReadGroups(root, "subject_groups");
            plan.BodyGroups = ReadGroups(root, "body_groups");
            plan.ReceivedFrom = ReadString(root, "received_from");
            plan.ReceivedThrough = ReadString(root, "received_through");
            plan.HasAttachments = ReadNullableBoolean(root, "has_attachments");
            plan.IsUnread = ReadNullableBoolean(root, "is_unread");
            plan.Scope = ReadString(root, "scope");
            plan.Normalize();
            ValidateScope(plan);
            ValidateDate(plan.ReceivedFrom, "开始日期");
            ValidateDate(plan.ReceivedThrough, "结束日期");
            ValidateDateOrder(plan);
            TrimList(plan.From);
            TrimList(plan.To);
            TrimList(plan.Cc);
            TrimGroups(plan.TextGroups);
            TrimGroups(plan.SubjectGroups);
            TrimGroups(plan.BodyGroups);
            return plan;
        }

        private static string ReadString(
            IDictionary<string, object> root,
            string key)
        {
            object value;
            if (!root.TryGetValue(key, out value) || value == null)
            {
                return string.Empty;
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static bool ReadBoolean(
            IDictionary<string, object> root,
            string key,
            bool defaultValue)
        {
            bool? value = ReadNullableBoolean(root, key);
            return value.HasValue ? value.Value : defaultValue;
        }

        private static bool? ReadNullableBoolean(
            IDictionary<string, object> root,
            string key)
        {
            object value;
            if (!root.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            return bool.TryParse(Convert.ToString(value), out parsed)
                ? (bool?)parsed
                : null;
        }

        private static List<string> ReadStringList(
            IDictionary<string, object> root,
            string key)
        {
            List<string> values = new List<string>();
            object raw;
            if (!root.TryGetValue(key, out raw) || raw == null)
            {
                return values;
            }

            object[] array = raw as object[];
            if (array == null)
            {
                return values;
            }

            foreach (object item in array)
            {
                if (item != null)
                {
                    values.Add(Convert.ToString(item));
                }
            }

            return values;
        }

        private static List<List<string>> ReadGroups(
            IDictionary<string, object> root,
            string key)
        {
            List<List<string>> groups = new List<List<string>>();
            object raw;
            if (!root.TryGetValue(key, out raw) || raw == null)
            {
                return groups;
            }

            object[] outer = raw as object[];
            if (outer == null)
            {
                return groups;
            }

            foreach (object rawGroup in outer)
            {
                object[] inner = rawGroup as object[];
                if (inner == null)
                {
                    continue;
                }

                List<string> group = new List<string>();
                foreach (object item in inner)
                {
                    if (item != null)
                    {
                        group.Add(Convert.ToString(item));
                    }
                }

                groups.Add(group);
            }

            return groups;
        }

        private static string StripMarkdownFence(string value)
        {
            string text = (value ?? string.Empty).Trim();
            if (!text.StartsWith("```", StringComparison.Ordinal))
            {
                return text;
            }

            int firstLineEnd = text.IndexOf('\n');
            int lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (firstLineEnd >= 0 && lastFence > firstLineEnd)
            {
                return text.Substring(
                    firstLineEnd + 1,
                    lastFence - firstLineEnd - 1).Trim();
            }

            return text;
        }

        private static void ValidateScope(SearchPlan plan)
        {
            if (!string.Equals(
                    plan.Scope,
                    "all_folders",
                    StringComparison.OrdinalIgnoreCase)
                && !string.Equals(
                    plan.Scope,
                    "current_folder",
                    StringComparison.OrdinalIgnoreCase))
            {
                plan.Scope = "all_folders";
            }
        }

        private static void ValidateDate(string value, string fieldName)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            DateTime parsed;
            if (!DateTime.TryParseExact(
                value,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out parsed))
            {
                throw new InvalidOperationException(
                    fieldName + "必须使用 YYYY-MM-DD 格式。");
            }
        }

        private static void ValidateDateOrder(SearchPlan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.ReceivedFrom)
                || string.IsNullOrWhiteSpace(plan.ReceivedThrough))
            {
                return;
            }

            DateTime from = DateTime.ParseExact(
                plan.ReceivedFrom,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture);
            DateTime through = DateTime.ParseExact(
                plan.ReceivedThrough,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture);
            if (from > through)
            {
                throw new InvalidOperationException("搜索开始日期不能晚于结束日期。");
            }
        }

        private static void TrimList(List<string> values)
        {
            for (int index = values.Count - 1; index >= 0; index--)
            {
                string value = (values[index] ?? string.Empty).Trim();
                if (value.Length == 0)
                {
                    values.RemoveAt(index);
                }
                else
                {
                    values[index] = value;
                }
            }

            if (values.Count > 20)
            {
                values.RemoveRange(20, values.Count - 20);
            }
        }

        private static void TrimGroups(List<List<string>> groups)
        {
            for (int index = groups.Count - 1; index >= 0; index--)
            {
                List<string> group = groups[index];
                if (group == null)
                {
                    groups.RemoveAt(index);
                    continue;
                }

                TrimList(group);
                if (group.Count == 0)
                {
                    groups.RemoveAt(index);
                }
            }

            if (groups.Count > 12)
            {
                groups.RemoveRange(12, groups.Count - 12);
            }
        }
    }
}
