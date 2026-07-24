using System.Collections.Generic;

namespace OutlookAiAssistant.Models
{
    /// <summary>
    /// Detached data for the selected mail and the separately stored Outlook
    /// messages that precede it in the same conversation.
    /// </summary>
    public sealed class ConversationSnapshot
    {
        public EmailSnapshot CurrentEmail { get; set; }
        public IList<EmailSnapshot> HistoryMessages { get; private set; }
        public int TotalConversationItems { get; set; }
        public bool ConversationAvailable { get; set; }
        public bool HistoryWasTruncated { get; set; }

        public ConversationSnapshot()
        {
            HistoryMessages = new List<EmailSnapshot>();
        }
    }
}
