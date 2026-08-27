using System;
using System.Collections.Generic;

namespace OutlookAiAssistant.Configuration
{
    internal static class DelimitedValues
    {
        public static IList<string> Parse(string value)
        {
            string[] parts = (value ?? string.Empty).Split(
                new[] { '\r', '\n', ',', ';', '，', '；' },
                StringSplitOptions.RemoveEmptyEntries);
            List<string> values = new List<string>();
            HashSet<string> seen = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < parts.Length; index++)
            {
                string part = parts[index].Trim();
                if (part.Length > 0 && seen.Add(part))
                {
                    values.Add(part);
                }
            }

            return values;
        }
    }
}
