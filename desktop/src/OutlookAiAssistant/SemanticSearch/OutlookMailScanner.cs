using System;
using System.Collections.Generic;
using System.Globalization;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.OutlookIntegration;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Enumerates Outlook mail metadata with MAPIFolder.GetTable() and reads a
    /// MailItem body only when the mail is new or changed.
    ///
    /// Threading: this class runs on the Outlook UI thread. It never hands a
    /// COM object to another thread and releases every temporary COM reference
    /// through ComRelease. Callers must invoke it from the UI thread.
    /// </summary>
    internal sealed class OutlookMailScanner
    {
        // MAPI property tags used to read metadata without opening MailItems.
        private const string StoreEntryIdProperty =
            "https://schemas.microsoft.com/mapi/proptag/0x0FFB0102";
        private const string LastModificationTimeProperty =
            "urn:schemas:httpmail:datereceived";
        private const string InternetMessageIdProperty =
            "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        private const string ConversationIdProperty =
            "http://schemas.microsoft.com/mapi/proptag/0x30130102";

        private readonly Outlook.Application _application;

        public OutlookMailScanner(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            _application = application;
        }

        /// <summary>
        /// Enumerates candidate mails for the semantic index. Only metadata is
        /// read; bodies are loaded later by ReadMailPayload.
        /// </summary>
        public IList<MailCandidate> EnumerateCandidates(
            AppSettings settings,
            Action<int> progress)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            List<MailCandidate> candidates = new List<MailCandidate>();
            DateTime cutoff = DateTime.Now.AddMonths(
                -Math.Max(1, settings.SemanticIndexMonths));
            int maximum = Math.Max(100, settings.SemanticIndexMaxMails);

            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            try
            {
                session = _application.Session;
                if (session == null)
                {
                    throw new InvalidOperationException("Outlook 会话尚未准备好。");
                }

                stores = session.Stores;
                if (stores == null)
                {
                    return candidates;
                }

                for (int storeIndex = 1;
                    storeIndex <= stores.Count && candidates.Count < maximum;
                    storeIndex++)
                {
                    Outlook.Store store = null;
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        store = stores[storeIndex];
                        if (store == null)
                        {
                            continue;
                        }

                        HashSet<string> excluded = GetExcludedFolderIds(store);
                        root = store.GetRootFolder();
                        if (root != null)
                        {
                            EnumerateFolder(
                                root,
                                excluded,
                                string.Empty,
                                cutoff,
                                maximum,
                                candidates,
                                progress);
                        }
                    }
                    catch
                    {
                        // One unavailable store must not stop other stores.
                    }
                    finally
                    {
                        ComRelease.Release(root);
                        ComRelease.Release(store);
                    }
                }
            }
            finally
            {
                ComRelease.Release(stores);
                ComRelease.Release(session);
            }

            return candidates;
        }

        private void EnumerateFolder(
            Outlook.MAPIFolder folder,
            ISet<string> excludedFolderIds,
            string relativePath,
            DateTime cutoff,
            int maximum,
            IList<MailCandidate> candidates,
            Action<int> progress)
        {
            string folderId = SafeString(delegate { return folder.EntryID; });
            if (folderId.Length > 0 && excludedFolderIds.Contains(folderId))
            {
                return;
            }

            bool isSent = IsSentFolder(folder);
            string folderPath = relativePath.Length == 0
                ? SafeString(delegate { return folder.Name; })
                : relativePath;

            ReadFolderTable(
                folder,
                folderId,
                folderPath,
                isSent,
                cutoff,
                maximum,
                candidates);

            if (progress != null)
            {
                progress(candidates.Count);
            }

            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                if (children == null)
                {
                    return;
                }

                for (int childIndex = 1;
                    childIndex <= children.Count && candidates.Count < maximum;
                    childIndex++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[childIndex];
                        if (child != null)
                        {
                            string childName = SafeString(
                                delegate { return child.Name; });
                            string childPath = folderPath.Length == 0
                                ? childName
                                : folderPath + "\\" + childName;
                            EnumerateFolder(
                                child,
                                excludedFolderIds,
                                childPath,
                                cutoff,
                                maximum,
                                candidates,
                                progress);
                        }
                    }
                    catch
                    {
                        // Continue with sibling folders.
                    }
                    finally
                    {
                        ComRelease.Release(child);
                    }
                }
            }
            finally
            {
                ComRelease.Release(children);
            }
        }

        private static void ReadFolderTable(
            Outlook.MAPIFolder folder,
            string folderId,
            string folderPath,
            bool isSentFolder,
            DateTime cutoff,
            int maximum,
            IList<MailCandidate> candidates)
        {
            Outlook.Table table = null;
            Outlook.Columns columns = null;
            try
            {
                table = folder.GetTable(
                    "[MessageClass] = 'IPM.Note'",
                    Outlook.OlTableContents.olUserItems);
                if (table == null)
                {
                    return;
                }

                columns = table.Columns;
                AddColumn(columns, "EntryID");
                AddColumn(columns, "Subject");
                AddColumn(columns, "SenderName");
                AddColumn(columns, "SenderEmailAddress");
                AddColumn(columns, "ReceivedTime");
                AddColumn(columns, "SentOn");
                AddColumn(columns, "LastModificationTime");
                AddColumn(columns, "MessageClass");
                AddColumn(columns, StoreEntryIdProperty);
                AddColumn(columns, InternetMessageIdProperty);
                AddColumn(columns, ConversationIdProperty);
                AddColumn(columns, LastModificationTimeProperty);

                table.Sort("[ReceivedTime]", true);

                while (!table.EndOfTable
                    && candidates.Count < maximum)
                {
                    Outlook.Row row = null;
                    try
                    {
                        row = table.GetNextRow();
                        if (row == null)
                        {
                            continue;
                        }

                        DateTime received = SafeRowDate(row, "ReceivedTime");
                        if (received != DateTime.MinValue && received < cutoff)
                        {
                            // Table is sorted descending by ReceivedTime, so
                            // once we pass the cutoff the rest are older.
                            break;
                        }

                        string entryId = SafeRowString(row, "EntryID");
                        if (string.IsNullOrWhiteSpace(entryId))
                        {
                            continue;
                        }

                        MailCandidate candidate = new MailCandidate();
                        candidate.EntryId = entryId;
                        candidate.StoreId = SafeRowBinary(row, StoreEntryIdProperty);
                        candidate.Subject = SafeRowString(row, "Subject");
                        candidate.SenderName = SafeRowString(row, "SenderName");
                        candidate.SenderEmail =
                            SafeRowString(row, "SenderEmailAddress");
                        candidate.ReceivedAt = received;
                        candidate.SentAt = SafeRowDate(row, "SentOn");
                        candidate.LastModifiedAt =
                            SafeRowDate(row, "LastModificationTime");
                        candidate.InternetMessageId =
                            SafeRowString(row, InternetMessageIdProperty);
                        candidate.ConversationId =
                            SafeRowBinary(row, ConversationIdProperty);
                        candidate.FolderId = folderId;
                        candidate.FolderPath = folderPath;
                        candidate.IsSentFolder = isSentFolder;
                        candidates.Add(candidate);
                    }
                    catch
                    {
                        // One exceptional row must not stop the scan.
                    }
                    finally
                    {
                        ComRelease.Release(row);
                    }
                }
            }
            catch
            {
                // Some folders do not support tables; skip them.
            }
            finally
            {
                ComRelease.Release(columns);
                ComRelease.Release(table);
            }
        }

        /// <summary>
        /// Reads the full MailItem for one candidate. This opens the item, so
        /// it is only called for mails that actually need (re)indexing.
        /// </summary>
        public MailIndexPayload ReadMailPayload(
            MailCandidate candidate,
            string ownerMailbox)
        {
            if (candidate == null)
            {
                throw new ArgumentNullException("candidate");
            }

            Outlook.NameSpace session = null;
            object item = null;
            try
            {
                session = _application.Session;
                item = string.IsNullOrWhiteSpace(candidate.StoreId)
                    ? session.GetItemFromID(candidate.EntryId, Type.Missing)
                    : session.GetItemFromID(
                        candidate.EntryId, candidate.StoreId);
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    return null;
                }

                MailIndexPayload payload = new MailIndexPayload();
                payload.EntryId = candidate.EntryId;
                payload.StoreId = candidate.StoreId;
                payload.Subject = FirstNonEmpty(
                    SafeString(delegate { return mail.Subject; }),
                    candidate.Subject);
                payload.SenderName = FirstNonEmpty(
                    SafeString(delegate { return mail.SenderName; }),
                    candidate.SenderName);
                payload.SenderEmail = FirstNonEmpty(
                    GetSenderEmail(mail), candidate.SenderEmail);
                payload.Recipients = SafeString(delegate { return mail.To; });
                payload.InternetMessageId = candidate.InternetMessageId;
                payload.ConversationId = candidate.ConversationId;
                payload.ReceivedAt = FormatDate(
                    SafeDate(
                        delegate { return mail.ReceivedTime; },
                        delegate { return candidate.ReceivedAt; }));
                payload.SentAt = FormatDate(
                    SafeDate(
                        delegate { return mail.SentOn; },
                        delegate { return candidate.SentAt; }));
                payload.LastModifiedAt = FormatDate(
                    SafeDate(
                        delegate { return mail.LastModificationTime; },
                        delegate { return candidate.LastModifiedAt; }));
                payload.FolderId = candidate.FolderId;
                payload.FolderPath = candidate.FolderPath;
                payload.OwnerMailbox = ownerMailbox ?? string.Empty;

                ReadAttachments(mail, payload);

                // Prefer the plain-text body; fall back to HTML when the
                // plain-text part is empty (common for HTML-only senders).
                payload.BodyText = SafeString(delegate { return mail.Body; });
                if (string.IsNullOrWhiteSpace(payload.BodyText))
                {
                    payload.HtmlBody =
                        SafeString(delegate { return mail.HTMLBody; });
                }

                return payload;
            }
            catch
            {
                // A moved, deleted or malformed mail is skipped and reported
                // by the caller as a single-item failure.
                return null;
            }
            finally
            {
                ComRelease.Release(item);
                ComRelease.Release(session);
            }
        }

        private static void ReadAttachments(
            Outlook.MailItem mail,
            MailIndexPayload payload)
        {
            Outlook.Attachments attachments = null;
            try
            {
                attachments = mail.Attachments;
                if (attachments == null)
                {
                    return;
                }

                payload.AttachmentCount = attachments.Count;
                List<string> names = new List<string>();
                for (int index = 1;
                    index <= attachments.Count && index <= 30;
                    index++)
                {
                    Outlook.Attachment attachment = null;
                    try
                    {
                        attachment = attachments[index];
                        string name = SafeString(
                            delegate { return attachment.FileName; });
                        if (!string.IsNullOrWhiteSpace(name))
                        {
                            names.Add(name);
                        }
                    }
                    catch
                    {
                        // Skip an unreadable attachment.
                    }
                    finally
                    {
                        ComRelease.Release(attachment);
                    }
                }

                payload.AttachmentNames = string.Join("; ", names.ToArray());
            }
            catch
            {
                payload.AttachmentCount = 0;
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
                    sender.Type, "EX", StringComparison.OrdinalIgnoreCase))
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

        public string GetOwnerMailbox()
        {
            Outlook.NameSpace session = null;
            Outlook.Recipient currentUser = null;
            try
            {
                session = _application.Session;
                currentUser = session == null ? null : session.CurrentUser;
                if (currentUser == null)
                {
                    return string.Empty;
                }

                return SafeString(delegate { return currentUser.Name; });
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                ComRelease.Release(currentUser);
                ComRelease.Release(session);
            }
        }

        private static bool IsSentFolder(Outlook.MAPIFolder folder)
        {
            return IsDefaultFolder(
                folder, Outlook.OlDefaultFolders.olFolderSentMail);
        }

        private static bool IsDefaultFolder(
            Outlook.MAPIFolder folder,
            Outlook.OlDefaultFolders type)
        {
            try
            {
                Outlook.Store store = folder.Store;
                if (store == null)
                {
                    return false;
                }

                Outlook.MAPIFolder sent = store.GetDefaultFolder(type);
                string sentId = sent == null
                    ? string.Empty
                    : SafeString(delegate { return sent.EntryID; });
                string folderId = SafeString(delegate { return folder.EntryID; });
                ComRelease.Release(sent);
                ComRelease.Release(store);
                return sentId.Length > 0
                    && string.Equals(sentId, folderId, StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
        }

        private static HashSet<string> GetExcludedFolderIds(Outlook.Store store)
        {
            HashSet<string> ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            Outlook.OlDefaultFolders[] types =
            {
                Outlook.OlDefaultFolders.olFolderDeletedItems,
                Outlook.OlDefaultFolders.olFolderJunk,
                Outlook.OlDefaultFolders.olFolderDrafts,
                Outlook.OlDefaultFolders.olFolderOutbox
            };
            for (int index = 0; index < types.Length; index++)
            {
                Outlook.MAPIFolder folder = null;
                try
                {
                    folder = store.GetDefaultFolder(types[index]);
                    string id = folder == null
                        ? string.Empty
                        : SafeString(delegate { return folder.EntryID; });
                    if (id.Length > 0)
                    {
                        ids.Add(id);
                    }
                }
                catch
                {
                    // PST and archive stores may not have every default folder.
                }
                finally
                {
                    ComRelease.Release(folder);
                }
            }

            return ids;
        }

        private static void AddColumn(Outlook.Columns columns, string name)
        {
            try
            {
                columns.Add(name);
            }
            catch
            {
                // Optional columns may be unsupported by some stores.
            }
        }

        private static string SafeRowString(Outlook.Row row, string column)
        {
            try
            {
                return Convert.ToString(row[column]) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string SafeRowBinary(Outlook.Row row, string column)
        {
            try
            {
                return row.BinaryToString(column) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime SafeRowDate(Outlook.Row row, string column)
        {
            try
            {
                object value = row[column];
                if (value == null)
                {
                    return DateTime.MinValue;
                }

                return Convert.ToDateTime(value);
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        private static string FormatDate(DateTime value)
        {
            if (value == DateTime.MinValue)
            {
                return string.Empty;
            }

            return value.ToUniversalTime().ToString(
                "yyyy-MM-ddTHH:mm:ssZ",
                CultureInfo.InvariantCulture);
        }

        private static string FirstNonEmpty(string primary, string fallback)
        {
            return string.IsNullOrWhiteSpace(primary) ? fallback : primary;
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
    }

    /// <summary>Metadata-only view of a candidate mail.</summary>
    internal sealed class MailCandidate
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public DateTime ReceivedAt { get; set; }
        public DateTime SentAt { get; set; }
        public DateTime LastModifiedAt { get; set; }
        public string InternetMessageId { get; set; }
        public string ConversationId { get; set; }
        public string FolderId { get; set; }
        public string FolderPath { get; set; }
        public bool IsSentFolder { get; set; }

        public MailCandidate()
        {
            EntryId = string.Empty;
            StoreId = string.Empty;
            Subject = string.Empty;
            SenderName = string.Empty;
            SenderEmail = string.Empty;
            InternetMessageId = string.Empty;
            ConversationId = string.Empty;
            FolderId = string.Empty;
            FolderPath = string.Empty;
        }
    }
}
