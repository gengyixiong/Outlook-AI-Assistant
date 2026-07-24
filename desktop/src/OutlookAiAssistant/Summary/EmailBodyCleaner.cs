using System;

namespace OutlookAiAssistant.Summary
{
    /// <summary>
    /// Removes only high-confidence quoted-reply separators from historical
    /// messages. The currently selected email is never cleaned because it may
    /// contain earlier correspondence that is not stored in this mailbox.
    /// </summary>
    public static class EmailBodyCleaner
    {
        private static readonly string[] QuoteMarkers =
        {
            "-----Original Message-----",
            "-----原始邮件-----",
            "-----原始消息-----",
            "________________________________"
        };

        private static readonly string[] HeaderMarkers =
        {
            "发件人:",
            "发件人：",
            "From:"
        };

        public static string RemoveQuotedHistory(string body)
        {
            string value = body ?? string.Empty;
            int boundary = value.Length;

            for (int index = 0; index < QuoteMarkers.Length; index++)
            {
                int candidate = FindMarkerAtLineStart(value, QuoteMarkers[index]);
                if (candidate >= 0 && candidate < boundary)
                {
                    boundary = candidate;
                }
            }

            int headerBoundary = FindOutlookHeaderBlock(value);
            if (headerBoundary >= 0 && headerBoundary < boundary)
            {
                boundary = headerBoundary;
            }

            return value.Substring(0, boundary).Trim();
        }

        private static int FindOutlookHeaderBlock(string value)
        {
            for (int index = 0; index < HeaderMarkers.Length; index++)
            {
                int searchFrom = 0;
                while (searchFrom < value.Length)
                {
                    int candidate = value.IndexOf(
                        HeaderMarkers[index],
                        searchFrom,
                        StringComparison.OrdinalIgnoreCase);
                    if (candidate < 0)
                    {
                        break;
                    }

                    bool lineStart = candidate == 0
                        || value[candidate - 1] == '\n'
                        || value[candidate - 1] == '\r';
                    int windowLength = Math.Min(800, value.Length - candidate);
                    string headerWindow = value.Substring(
                        candidate,
                        windowLength);
                    if (lineStart && CountHeaderSignals(headerWindow) >= 2)
                    {
                        return candidate;
                    }

                    searchFrom = candidate + HeaderMarkers[index].Length;
                }
            }

            return -1;
        }

        private static int CountHeaderSignals(string value)
        {
            int count = 0;
            if (ContainsAny(
                value,
                "\nSent:",
                "\n发送时间:",
                "\n发送时间：",
                "\n日期:",
                "\n日期："))
            {
                count++;
            }

            if (ContainsAny(
                value,
                "\nTo:",
                "\n收件人:",
                "\n收件人："))
            {
                count++;
            }

            if (ContainsAny(
                value,
                "\nSubject:",
                "\n主题:",
                "\n主题："))
            {
                count++;
            }

            return count;
        }

        private static bool ContainsAny(
            string value,
            params string[] candidates)
        {
            for (int index = 0; index < candidates.Length; index++)
            {
                if (value.IndexOf(
                    candidates[index],
                    StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static int FindMarkerAtLineStart(string value, string marker)
        {
            int searchFrom = 0;
            while (searchFrom < value.Length)
            {
                int candidate = value.IndexOf(
                    marker,
                    searchFrom,
                    StringComparison.OrdinalIgnoreCase);
                if (candidate < 0)
                {
                    return -1;
                }

                bool lineStart = candidate == 0
                    || value[candidate - 1] == '\n'
                    || value[candidate - 1] == '\r';
                if (lineStart)
                {
                    return candidate;
                }

                searchFrom = candidate + marker.Length;
            }

            return -1;
        }
    }
}
