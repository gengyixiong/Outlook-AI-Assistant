using System;

namespace OutlookAiAssistant.Models
{
    /// <summary>
    /// A detached, plain-data copy of the currently selected Outlook message.
    /// No COM object is kept after the snapshot is created.
    /// </summary>
    public sealed class EmailSnapshot
    {
        public string EntryId { get; set; }
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string To { get; set; }
        public string Cc { get; set; }
        public DateTime ReceivedAt { get; set; }
        public string Body { get; set; }
        public int AttachmentCount { get; set; }
        public string CurrentUserName { get; set; }
        public string CurrentUserEmail { get; set; }

        public EmailSnapshot()
        {
            EntryId = string.Empty;
            Subject = string.Empty;
            SenderName = string.Empty;
            SenderEmail = string.Empty;
            To = string.Empty;
            Cc = string.Empty;
            Body = string.Empty;
            CurrentUserName = string.Empty;
            CurrentUserEmail = string.Empty;
        }
    }
}
