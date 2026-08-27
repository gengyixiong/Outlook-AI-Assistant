using System;
using System.Collections.Generic;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.OutlookIntegration;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.EntityIndex
{
    /// <summary>
    /// Discovers senders with lightweight MailItem fields first, then loads
    /// bodies only for the earliest and three most recent representative mails.
    /// </summary>
    internal sealed class OutlookContactScanner
    {
        private readonly Outlook.Application _application;

        public OutlookContactScanner(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            _application = application;
        }

        public ContactDiscoveryResult Discover(AppSettings settings)
        {
            Dictionary<string, ContactCandidate> candidates =
                new Dictionary<string, ContactCandidate>(
                    StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> smtpCache =
                new Dictionary<string, string>(
                    StringComparer.OrdinalIgnoreCase);
            HashSet<string> selfEmails = ReadConfiguredSelfEmails(settings);
            List<FolderContextEntry> folderContext =
                new List<FolderContextEntry>();
            Outlook.NameSpace session = null;
            Outlook.Stores stores = null;
            try
            {
                session = _application.Session;
                if (session == null)
                {
                    throw new InvalidOperationException(
                        "Outlook 会话尚未准备好。");
                }

                AddCurrentUserEmail(session, selfEmails);
                stores = session.Stores;
                if (stores == null)
                {
                    return new ContactDiscoveryResult();
                }

                for (int storeIndex = 1;
                    storeIndex <= stores.Count;
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
                        HashSet<string> standard = GetStandardFolderIds(store);
                        root = store.GetRootFolder();
                        if (root != null)
                        {
                            string rootName = SafeString(
                                delegate { return store.DisplayName; });
                            if (rootName.Length == 0)
                            {
                                rootName = SafeString(
                                    delegate { return root.Name; });
                            }

                            ScanFolder(
                                root,
                                excluded,
                                standard,
                                selfEmails,
                                smtpCache,
                                candidates,
                                rootName,
                                string.Empty,
                                folderContext);
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

            List<ContactCandidate> result =
                new List<ContactCandidate>(candidates.Values);
            result.Sort(delegate(ContactCandidate left, ContactCandidate right)
            {
                return string.Compare(
                    left.Email,
                    right.Email,
                    StringComparison.OrdinalIgnoreCase);
            });
            folderContext.Sort(delegate(
                FolderContextEntry left,
                FolderContextEntry right)
            {
                int rootComparison = string.Compare(
                    left.Root,
                    right.Root,
                    StringComparison.OrdinalIgnoreCase);
                return rootComparison != 0
                    ? rootComparison
                    : string.Compare(
                        left.RelativePath,
                        right.RelativePath,
                        StringComparison.OrdinalIgnoreCase);
            });
            return new ContactDiscoveryResult
            {
                Candidates = result,
                FolderContext = folderContext
            };
        }

        public IList<ContactRepresentativeMessage> LoadRepresentativeMessages(
            ContactCandidate candidate)
        {
            List<ContactRepresentativeMessage> messages =
                new List<ContactRepresentativeMessage>();
            Outlook.NameSpace session = null;
            try
            {
                session = _application.Session;
                if (session == null)
                {
                    return messages;
                }

                IList<ContactMessageReference> references =
                    candidate.GetRepresentativeReferences();
                for (int index = 0; index < references.Count; index++)
                {
                    object item = null;
                    try
                    {
                        ContactMessageReference reference = references[index];
                        item = string.IsNullOrWhiteSpace(reference.StoreId)
                            ? session.GetItemFromID(reference.EntryId, Type.Missing)
                            : session.GetItemFromID(
                                reference.EntryId,
                                reference.StoreId);
                        Outlook.MailItem mail = item as Outlook.MailItem;
                        if (mail == null)
                        {
                            continue;
                        }

                        messages.Add(new ContactRepresentativeMessage
                        {
                            Subject = SafeString(delegate { return mail.Subject; }),
                            SenderName = SafeString(
                                delegate { return mail.SenderName; }),
                            ReceivedAt = SafeDate(
                                delegate { return mail.ReceivedTime; }),
                            Body = SafeString(delegate { return mail.Body; })
                        });
                    }
                    catch
                    {
                        // A moved or malformed representative mail is skipped.
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

            messages.Sort(delegate(
                ContactRepresentativeMessage left,
                ContactRepresentativeMessage right)
            {
                return left.ReceivedAt.CompareTo(right.ReceivedAt);
            });
            return messages;
        }

        private static void ScanFolder(
            Outlook.MAPIFolder folder,
            ISet<string> excludedFolderIds,
            ISet<string> standardFolderIds,
            ISet<string> selfEmails,
            IDictionary<string, string> smtpCache,
            IDictionary<string, ContactCandidate> candidates,
            string rootName,
            string relativePath,
            ICollection<FolderContextEntry> folderEntries)
        {
            string folderId = SafeString(delegate { return folder.EntryID; });
            if (folderId.Length > 0 && excludedFolderIds.Contains(folderId))
            {
                return;
            }

            FolderContextEntry folderEntry = null;
            if (relativePath.Length > 0
                && !standardFolderIds.Contains(folderId)
                && IsMailFolder(folder))
            {
                folderEntry = new FolderContextEntry
                {
                    Root = rootName,
                    RelativePath = relativePath
                };
                folderEntries.Add(folderEntry);
            }

            ScanFolderItems(
                folder,
                selfEmails,
                smtpCache,
                candidates,
                folderEntry);

            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                if (children == null)
                {
                    return;
                }

                for (int childIndex = 1;
                    childIndex <= children.Count;
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
                            string childPath = relativePath.Length == 0
                                ? childName
                                : relativePath + "\\" + childName;
                            ScanFolder(
                                child,
                                excludedFolderIds,
                                standardFolderIds,
                                selfEmails,
                                smtpCache,
                                candidates,
                                rootName,
                                childPath,
                                folderEntries);
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

        private static void ScanFolderItems(
            Outlook.MAPIFolder folder,
            ISet<string> selfEmails,
            IDictionary<string, string> smtpCache,
            IDictionary<string, ContactCandidate> candidates,
            FolderContextEntry folderEntry)
        {
            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                if (items == null)
                {
                    return;
                }

                string storeId = SafeString(delegate { return folder.StoreID; });
                for (int itemIndex = 1;
                    itemIndex <= items.Count;
                    itemIndex++)
                {
                    object item = null;
                    try
                    {
                        item = items[itemIndex];
                        Outlook.MailItem mail = item as Outlook.MailItem;
                        if (mail == null)
                        {
                            continue;
                        }

                        string email = GetSenderEmail(mail, smtpCache)
                            .Trim()
                            .ToLowerInvariant();
                        if (email.IndexOf('@') <= 0
                            || selfEmails.Contains(email))
                        {
                            continue;
                        }

                        ContactCandidate candidate;
                        if (!candidates.TryGetValue(email, out candidate))
                        {
                            // ponytail: a hard ceiling bounds manual scan/API
                            // cost; raise it if a real mailbox exceeds 1000 senders.
                            if (candidates.Count >= 1000)
                            {
                                continue;
                            }

                            candidate = new ContactCandidate(email);
                            candidates[email] = candidate;
                        }

                        candidate.Add(
                            SafeString(delegate { return mail.SenderName; }),
                            new ContactMessageReference
                            {
                                EntryId = SafeString(
                                    delegate { return mail.EntryID; }),
                                StoreId = storeId,
                                ReceivedAt = SafeDate(
                                    delegate { return mail.ReceivedTime; })
                            });
                        if (folderEntry != null)
                        {
                            string contactId = ContactProfile.CreateId(email);
                            AddUnique(folderEntry.ContactIds, contactId);
                            candidate.AddFolderLabel(
                                folderEntry.Root + " / "
                                    + folderEntry.RelativePath);
                        }
                    }
                    catch
                    {
                        // One exceptional item must not fail the index.
                    }
                    finally
                    {
                        ComRelease.Release(item);
                    }
                }
            }
            catch
            {
                // Some Outlook folders do not expose an Items collection.
            }
            finally
            {
                ComRelease.Release(items);
            }
        }

        private static HashSet<string> GetExcludedFolderIds(
            Outlook.Store store)
        {
            HashSet<string> ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            Outlook.OlDefaultFolders[] types =
            {
                Outlook.OlDefaultFolders.olFolderDeletedItems,
                Outlook.OlDefaultFolders.olFolderJunk,
                Outlook.OlDefaultFolders.olFolderDrafts,
                Outlook.OlDefaultFolders.olFolderOutbox,
                Outlook.OlDefaultFolders.olFolderSentMail
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

        private static void AddDefaultFolderId(
            Outlook.Store store,
            Outlook.OlDefaultFolders folderType,
            ISet<string> ids)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = store.GetDefaultFolder(folderType);
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
                // A store does not necessarily expose every default folder.
            }
            finally
            {
                ComRelease.Release(folder);
            }
        }

        private static HashSet<string> GetStandardFolderIds(
            Outlook.Store store)
        {
            HashSet<string> ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            Array folderTypes = Enum.GetValues(
                typeof(Outlook.OlDefaultFolders));
            foreach (object value in folderTypes)
            {
                AddDefaultFolderId(
                    store,
                    (Outlook.OlDefaultFolders)value,
                    ids);
            }

            Outlook.Folders searchFolders = null;
            try
            {
                searchFolders = store.GetSearchFolders();
                if (searchFolders != null)
                {
                    for (int index = 1; index <= searchFolders.Count; index++)
                    {
                        Outlook.MAPIFolder folder = null;
                        try
                        {
                            folder = searchFolders[index];
                            string id = folder == null
                                ? string.Empty
                                : SafeString(
                                    delegate { return folder.EntryID; });
                            if (id.Length > 0)
                            {
                                ids.Add(id);
                            }
                        }
                        finally
                        {
                            ComRelease.Release(folder);
                        }
                    }
                }
            }
            catch
            {
                // Some stores do not support Search Folders.
            }
            finally
            {
                ComRelease.Release(searchFolders);
            }

            return ids;
        }

        private static bool IsMailFolder(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.DefaultItemType == Outlook.OlItemType.olMailItem;
            }
            catch
            {
                return false;
            }
        }

        private static void AddUnique(ICollection<string> values, string value)
        {
            foreach (string existing in values)
            {
                if (string.Equals(
                    existing,
                    value,
                    StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            values.Add(value);
        }

        private static HashSet<string> ReadConfiguredSelfEmails(
            AppSettings settings)
        {
            HashSet<string> emails = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            IList<string> values = DelimitedValues.Parse(
                settings == null
                    ? string.Empty
                    : settings.IdentityEmailAddresses);
            for (int index = 0; index < values.Count; index++)
            {
                emails.Add(values[index].Trim());
            }

            return emails;
        }

        private static void AddCurrentUserEmail(
            Outlook.NameSpace session,
            ISet<string> selfEmails)
        {
            Outlook.Recipient currentUser = null;
            Outlook.AddressEntry addressEntry = null;
            Outlook.ExchangeUser exchangeUser = null;
            try
            {
                currentUser = session.CurrentUser;
                addressEntry = currentUser == null
                    ? null
                    : currentUser.AddressEntry;
                if (addressEntry == null)
                {
                    return;
                }

                string email = string.Empty;
                if (string.Equals(
                    addressEntry.Type,
                    "EX",
                    StringComparison.OrdinalIgnoreCase))
                {
                    exchangeUser = addressEntry.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        email = exchangeUser.PrimarySmtpAddress;
                    }
                }
                else
                {
                    email = addressEntry.Address;
                }

                if (!string.IsNullOrWhiteSpace(email))
                {
                    selfEmails.Add(email.Trim());
                }
            }
            catch
            {
                // Self filtering is helpful but not required.
            }
            finally
            {
                ComRelease.Release(exchangeUser);
                ComRelease.Release(addressEntry);
                ComRelease.Release(currentUser);
            }
        }

        private static string GetSenderEmail(
            Outlook.MailItem mail,
            IDictionary<string, string> smtpCache)
        {
            string fallback = SafeString(
                delegate { return mail.SenderEmailAddress; });
            if (fallback.IndexOf('@') > 0)
            {
                return fallback;
            }

            string cached;
            if (fallback.Length > 0
                && smtpCache.TryGetValue(fallback, out cached))
            {
                return cached;
            }

            Outlook.AddressEntry sender = null;
            Outlook.ExchangeUser exchangeUser = null;
            string email = fallback;
            try
            {
                sender = mail.Sender;
                if (sender != null
                    && string.Equals(
                        sender.Type,
                        "EX",
                        StringComparison.OrdinalIgnoreCase))
                {
                    exchangeUser = sender.GetExchangeUser();
                    if (exchangeUser != null
                        && !string.IsNullOrWhiteSpace(
                            exchangeUser.PrimarySmtpAddress))
                    {
                        email = exchangeUser.PrimarySmtpAddress;
                    }
                }
                else if (sender != null
                    && !string.IsNullOrWhiteSpace(sender.Address))
                {
                    email = sender.Address;
                }
            }
            catch
            {
                email = fallback;
            }
            finally
            {
                ComRelease.Release(exchangeUser);
                ComRelease.Release(sender);
            }

            if (fallback.Length > 0)
            {
                smtpCache[fallback] = email ?? string.Empty;
            }

            return email ?? string.Empty;
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

        private static DateTime SafeDate(Func<DateTime> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return DateTime.MinValue;
            }
        }
    }

    internal sealed class ContactCandidate
    {
        private ContactMessageReference _earliest;
        private readonly List<ContactMessageReference> _recent;
        private readonly HashSet<string> _entryIds;
        private readonly HashSet<string> _folderLabelSet;
        private DateTime _displayNameDate;

        public string Email { get; private set; }
        public string DisplayName { get; private set; }
        public IList<string> FolderLabels { get; private set; }

        public ContactCandidate(string email)
        {
            Email = email ?? string.Empty;
            DisplayName = string.Empty;
            _displayNameDate = DateTime.MinValue;
            _recent = new List<ContactMessageReference>();
            _entryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _folderLabelSet = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            FolderLabels = new List<string>();
        }

        public void AddFolderLabel(string value)
        {
            string label = (value ?? string.Empty).Trim();
            // ponytail: 20 labels bound prompt size; raise only if real users
            // routinely classify one contact into more folders.
            if (label.Length > 0
                && FolderLabels.Count < 20
                && _folderLabelSet.Add(label))
            {
                FolderLabels.Add(label);
            }
        }

        public void Add(string displayName, ContactMessageReference reference)
        {
            if (reference == null || string.IsNullOrWhiteSpace(reference.EntryId))
            {
                return;
            }

            if (!_entryIds.Add(reference.EntryId))
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(displayName)
                && reference.ReceivedAt >= _displayNameDate)
            {
                DisplayName = displayName.Trim();
                _displayNameDate = reference.ReceivedAt;
            }

            if (_earliest == null
                || reference.ReceivedAt < _earliest.ReceivedAt)
            {
                _earliest = reference;
            }

            _recent.Add(reference);
            _recent.Sort(delegate(
                ContactMessageReference left,
                ContactMessageReference right)
            {
                return right.ReceivedAt.CompareTo(left.ReceivedAt);
            });
            if (_recent.Count > 3)
            {
                _recent.RemoveAt(_recent.Count - 1);
            }
        }

        public IList<ContactMessageReference> GetRepresentativeReferences()
        {
            List<ContactMessageReference> result =
                new List<ContactMessageReference>();
            HashSet<string> entryIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            if (_earliest != null && entryIds.Add(_earliest.EntryId))
            {
                result.Add(_earliest);
            }

            for (int index = _recent.Count - 1; index >= 0; index--)
            {
                if (entryIds.Add(_recent[index].EntryId))
                {
                    result.Add(_recent[index]);
                }
            }

            return result;
        }
    }

    internal sealed class ContactDiscoveryResult
    {
        public IList<ContactCandidate> Candidates { get; set; }
        public IList<FolderContextEntry> FolderContext { get; set; }

        public ContactDiscoveryResult()
        {
            Candidates = new List<ContactCandidate>();
            FolderContext = new List<FolderContextEntry>();
        }
    }

    internal sealed class ContactMessageReference
    {
        public string EntryId { get; set; }
        public string StoreId { get; set; }
        public DateTime ReceivedAt { get; set; }

        public ContactMessageReference()
        {
            EntryId = string.Empty;
            StoreId = string.Empty;
        }
    }

    internal sealed class ContactRepresentativeMessage
    {
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public DateTime ReceivedAt { get; set; }
        public string Body { get; set; }

        public ContactRepresentativeMessage()
        {
            Subject = string.Empty;
            SenderName = string.Empty;
            Body = string.Empty;
        }
    }
}
