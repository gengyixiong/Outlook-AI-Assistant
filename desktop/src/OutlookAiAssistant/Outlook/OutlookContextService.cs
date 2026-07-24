using System;
using System.Collections.Generic;
using OutlookAiAssistant.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.OutlookIntegration
{
    /// <summary>
    /// Reads Outlook data only when explicitly called by the summary button.
    /// It does not subscribe to incoming-mail or selection-change events.
    /// </summary>
    public sealed class OutlookContextService
    {
        private const string StoreEntryIdProperty =
            "https://schemas.microsoft.com/mapi/proptag/0x0FFB0102";

        public const int DefaultMaximumConversationMessages = 50;

        private readonly Outlook.Application _application;

        public OutlookContextService(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            _application = application;
        }

        public EmailSnapshot GetCurrentEmail()
        {
            return GetCurrentConversation(1).CurrentEmail;
        }

        public ConversationSnapshot GetCurrentConversation(
            int maximumMessages)
        {
            if (maximumMessages < 1)
            {
                throw new ArgumentOutOfRangeException("maximumMessages");
            }

            ConversationSnapshot snapshot = TryReadInspectorConversation(
                maximumMessages);
            if (snapshot != null)
            {
                return snapshot;
            }

            snapshot = TryReadExplorerConversation(maximumMessages);
            if (snapshot != null)
            {
                return snapshot;
            }

            throw new InvalidOperationException(
                "没有找到当前邮件。请先在 Outlook 中选中或打开一封邮件。");
        }

        private ConversationSnapshot TryReadInspectorConversation(
            int maximumMessages)
        {
            Outlook.Inspector inspector = null;
            object currentItem = null;
            try
            {
                inspector = _application.ActiveInspector();
                if (inspector == null)
                {
                    return null;
                }

                currentItem = inspector.CurrentItem;
                Outlook.MailItem mail = currentItem as Outlook.MailItem;
                return mail == null
                    ? null
                    : CreateConversationSnapshot(mail, maximumMessages);
            }
            finally
            {
                ComRelease.Release(currentItem);
                ComRelease.Release(inspector);
            }
        }

        private ConversationSnapshot TryReadExplorerConversation(
            int maximumMessages)
        {
            Outlook.Explorer explorer = null;
            Outlook.Selection selection = null;
            object selectedItem = null;
            try
            {
                explorer = _application.ActiveExplorer();
                if (explorer == null)
                {
                    return null;
                }

                selection = explorer.Selection;
                if (selection == null || selection.Count < 1)
                {
                    return null;
                }

                selectedItem = selection[1];
                Outlook.MailItem mail = selectedItem as Outlook.MailItem;
                return mail == null
                    ? null
                    : CreateConversationSnapshot(mail, maximumMessages);
            }
            finally
            {
                ComRelease.Release(selectedItem);
                ComRelease.Release(selection);
                ComRelease.Release(explorer);
            }
        }

        private ConversationSnapshot CreateConversationSnapshot(
            Outlook.MailItem currentMail,
            int maximumMessages)
        {
            string currentUserName;
            string currentUserEmail;
            GetCurrentUser(out currentUserName, out currentUserEmail);

            ConversationSnapshot result = new ConversationSnapshot();
            result.CurrentEmail = CreateSnapshot(
                currentMail,
                currentUserName,
                currentUserEmail);
            result.TotalConversationItems = 1;

            Outlook.Conversation conversation = null;
            Outlook.Table table = null;
            Outlook.Columns columns = null;
            Outlook.Column storeColumn = null;
            try
            {
                conversation = currentMail.GetConversation();
                if (conversation == null)
                {
                    return result;
                }

                table = conversation.GetTable();
                if (table == null)
                {
                    return result;
                }

                result.ConversationAvailable = true;
                result.TotalConversationItems = Math.Max(
                    1,
                    table.GetRowCount());
                columns = table.Columns;
                storeColumn = columns.Add(StoreEntryIdProperty);

                List<ConversationItemReference> references =
                    ReadConversationReferences(
                        table,
                        result.CurrentEmail);
                bool truncated;
                IList<ConversationItemReference> selected =
                    SelectReferences(
                        references,
                        result.CurrentEmail.EntryId,
                        maximumMessages - 1,
                        out truncated);
                result.HistoryWasTruncated = truncated;
                LoadHistoryMessages(
                    selected,
                    result,
                    currentUserName,
                    currentUserEmail);
                SortChronologically(result.HistoryMessages);
                return result;
            }
            catch
            {
                // Conversation view can be disabled or unsupported by a store.
                // The current mail remains available and can contain quoted
                // history, so summary generation should still proceed.
                result.ConversationAvailable = false;
                result.HistoryMessages.Clear();
                result.TotalConversationItems = 1;
                result.HistoryWasTruncated = false;
                return result;
            }
            finally
            {
                ComRelease.Release(storeColumn);
                ComRelease.Release(columns);
                ComRelease.Release(table);
                ComRelease.Release(conversation);
            }
        }

        private static List<ConversationItemReference>
            ReadConversationReferences(
                Outlook.Table table,
                EmailSnapshot current)
        {
            List<ConversationItemReference> references =
                new List<ConversationItemReference>();
            while (!table.EndOfTable)
            {
                Outlook.Row row = null;
                try
                {
                    row = table.GetNextRow();
                    if (row == null)
                    {
                        continue;
                    }

                    string messageClass = Convert.ToString(
                        row["MessageClass"]);
                    if (!messageClass.StartsWith(
                        "IPM.Note",
                        StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    DateTime sortTime = SafeRowDate(row["CreationTime"]);
                    if (current.ReceivedAt != DateTime.MinValue
                        && sortTime != DateTime.MinValue
                        && sortTime > current.ReceivedAt.AddMinutes(5))
                    {
                        continue;
                    }

                    ConversationItemReference reference =
                        new ConversationItemReference();
                    reference.EntryId = Convert.ToString(row["EntryID"]);
                    reference.StoreId = SafeStoreId(row);
                    reference.SortTime = sortTime;
                    if (!string.IsNullOrWhiteSpace(reference.EntryId))
                    {
                        references.Add(reference);
                    }
                }
                finally
                {
                    ComRelease.Release(row);
                }
            }

            references.Sort(delegate(
                ConversationItemReference left,
                ConversationItemReference right)
            {
                return left.SortTime.CompareTo(right.SortTime);
            });
            return references;
        }

        private static IList<ConversationItemReference> SelectReferences(
            IList<ConversationItemReference> references,
            string currentEntryId,
            int maximumHistoryMessages,
            out bool truncated)
        {
            List<ConversationItemReference> unique =
                new List<ConversationItemReference>();
            HashSet<string> entryIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < references.Count; index++)
            {
                ConversationItemReference reference = references[index];
                if (string.Equals(
                    reference.EntryId,
                    currentEntryId,
                    StringComparison.OrdinalIgnoreCase)
                    || !entryIds.Add(reference.EntryId))
                {
                    continue;
                }

                unique.Add(reference);
            }

            int safeMaximum = Math.Max(0, maximumHistoryMessages);
            truncated = unique.Count > safeMaximum;
            if (!truncated)
            {
                return unique;
            }

            List<ConversationItemReference> selected =
                new List<ConversationItemReference>();
            if (safeMaximum == 0)
            {
                return selected;
            }

            int earliestCount = Math.Min(
                10,
                Math.Max(1, safeMaximum / 4));
            earliestCount = Math.Min(earliestCount, safeMaximum);
            for (int index = 0; index < earliestCount; index++)
            {
                selected.Add(unique[index]);
            }

            int recentCount = safeMaximum - earliestCount;
            int recentStart = unique.Count - recentCount;
            for (int index = recentStart; index < unique.Count; index++)
            {
                if (index >= earliestCount)
                {
                    selected.Add(unique[index]);
                }
            }

            return selected;
        }

        private void LoadHistoryMessages(
            IList<ConversationItemReference> references,
            ConversationSnapshot result,
            string currentUserName,
            string currentUserEmail)
        {
            Outlook.NameSpace session = null;
            try
            {
                session = _application.Session;
                if (session == null)
                {
                    return;
                }

                for (int index = 0; index < references.Count; index++)
                {
                    object item = null;
                    try
                    {
                        ConversationItemReference reference =
                            references[index];
                        item = string.IsNullOrWhiteSpace(reference.StoreId)
                            ? session.GetItemFromID(
                                reference.EntryId,
                                Type.Missing)
                            : session.GetItemFromID(
                                reference.EntryId,
                                reference.StoreId);
                        Outlook.MailItem mail = item as Outlook.MailItem;
                        if (mail != null)
                        {
                            result.HistoryMessages.Add(
                                CreateSnapshot(
                                    mail,
                                    currentUserName,
                                    currentUserEmail));
                        }
                    }
                    catch
                    {
                        // One moved or inaccessible message must not prevent
                        // the rest of the conversation from being summarized.
                    }
                    finally
                    {
                        ComRelease.Release(item);
                    }
                }
            }
            finally
            {
                ComRelease.Release(session);
            }
        }

        private static void SortChronologically(
            IList<EmailSnapshot> messages)
        {
            List<EmailSnapshot> list = messages as List<EmailSnapshot>;
            if (list == null)
            {
                return;
            }

            list.Sort(delegate(EmailSnapshot left, EmailSnapshot right)
            {
                return left.ReceivedAt.CompareTo(right.ReceivedAt);
            });
        }

        private EmailSnapshot CreateSnapshot(
            Outlook.MailItem mail,
            string currentUserName,
            string currentUserEmail)
        {
            EmailSnapshot snapshot = new EmailSnapshot();
            snapshot.EntryId = SafeString(delegate { return mail.EntryID; });
            snapshot.Subject = SafeString(delegate { return mail.Subject; });
            snapshot.SenderName = SafeString(delegate { return mail.SenderName; });
            snapshot.SenderEmail = GetSenderEmail(mail);
            snapshot.To = SafeString(delegate { return mail.To; });
            snapshot.Cc = SafeString(delegate { return mail.CC; });
            snapshot.Body = SafeString(delegate { return mail.Body; });
            snapshot.ReceivedAt = SafeDate(
                delegate { return mail.ReceivedTime; },
                delegate { return mail.SentOn; });
            snapshot.AttachmentCount = GetAttachmentCount(mail);
            snapshot.CurrentUserName = currentUserName;
            snapshot.CurrentUserEmail = currentUserEmail;
            return snapshot;
        }

        private static int GetAttachmentCount(Outlook.MailItem mail)
        {
            Outlook.Attachments attachments = null;
            try
            {
                attachments = mail.Attachments;
                return attachments == null ? 0 : attachments.Count;
            }
            catch
            {
                return 0;
            }
            finally
            {
                ComRelease.Release(attachments);
            }
        }

        private static string GetSenderEmail(Outlook.MailItem mail)
        {
            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchangeUser = null;
            try
            {
                string fallback = SafeString(
                    delegate { return mail.SenderEmailAddress; });
                sender = mail.Sender;
                if (sender == null)
                {
                    return fallback;
                }

                if (string.Equals(
                    sender.Type,
                    "EX",
                    StringComparison.OrdinalIgnoreCase))
                {
                    exchangeUser = sender.GetExchangeUser();
                    if (exchangeUser != null
                        && !string.IsNullOrWhiteSpace(
                            exchangeUser.PrimarySmtpAddress))
                    {
                        return exchangeUser.PrimarySmtpAddress;
                    }
                }

                return string.IsNullOrWhiteSpace(sender.Address)
                    ? fallback
                    : sender.Address;
            }
            catch
            {
                return SafeString(
                    delegate { return mail.SenderEmailAddress; });
            }
            finally
            {
                ComRelease.Release(exchangeUser);
                ComRelease.Release(sender);
            }
        }

        private void GetCurrentUser(
            out string displayName,
            out string email)
        {
            displayName = string.Empty;
            email = string.Empty;
            Outlook.NameSpace session = null;
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry addressEntry = null;
            Outlook.ExchangeUser exchangeUser = null;

            try
            {
                session = _application.Session;
                currentUser = session == null ? null : session.CurrentUser;
                if (currentUser == null)
                {
                    return;
                }

                displayName = currentUser.Name ?? string.Empty;
                addressEntry = currentUser.AddressEntry;
                if (addressEntry == null)
                {
                    return;
                }

                if (string.Equals(
                    addressEntry.Type,
                    "EX",
                    StringComparison.OrdinalIgnoreCase))
                {
                    exchangeUser = addressEntry.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        email =
                            exchangeUser.PrimarySmtpAddress ?? string.Empty;
                    }
                }
                else
                {
                    email = addressEntry.Address ?? string.Empty;
                }
            }
            catch
            {
                // Identity helps personalization but is not required.
            }
            finally
            {
                ComRelease.Release(exchangeUser);
                ComRelease.Release(addressEntry);
                ComRelease.Release(currentUser);
                ComRelease.Release(session);
            }
        }

        private static string SafeStoreId(Outlook.Row row)
        {
            try
            {
                return row.BinaryToString(StoreEntryIdProperty)
                    ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime SafeRowDate(object value)
        {
            try
            {
                return Convert.ToDateTime(value);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        private static string SafeString(Func<string> getter)
        {
            try
            {
                return getter() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime SafeDate(
            Func<DateTime> primary,
            Func<DateTime> fallback)
        {
            try
            {
                return primary();
            }
            catch
            {
                try
                {
                    return fallback();
                }
                catch
                {
                    return DateTime.MinValue;
                }
            }
        }

        private sealed class ConversationItemReference
        {
            public string EntryId { get; set; }
            public string StoreId { get; set; }
            public DateTime SortTime { get; set; }

            public ConversationItemReference()
            {
                EntryId = string.Empty;
                StoreId = string.Empty;
            }
        }
    }
}
