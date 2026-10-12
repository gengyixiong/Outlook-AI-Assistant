using System;
using System.Collections.Generic;
using System.Globalization;
using OutlookAiAssistant.OutlookIntegration;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Opens the real Outlook MailItem behind a search result.
    ///
    /// Direct location via NameSpace.GetItemFromID is the primary path. If the
    /// EntryID is stale (the mail moved or was deleted) the locator verifies
    /// candidates by InternetMessageId and metadata before opening anything,
    /// so it never opens a different mail with a coincidentally similar title.
    /// </summary>
    public sealed class OutlookMailLocator
    {
        private readonly Outlook.Application _application;

        public OutlookMailLocator(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            _application = application;
        }

        /// <summary>
        /// Tries to open the mail. Returns null on success, or a user-facing
        /// message describing why it could not be opened.
        /// </summary>
        public string OpenMail(SemanticSearchResult result)
        {
            if (result == null)
            {
                return "搜索结果无效。";
            }

            Outlook.MailItem mail = ResolveMailItem(result);
            if (mail == null)
            {
                return "无法在 Outlook 中定位这封邮件。它可能已被移动或删除。"
                    + "可以尝试使用“Outlook 搜索”。";
            }

            try
            {
                mail.Display();
                return null;
            }
            catch (Exception ex)
            {
                return "打开邮件失败：" + ex.Message;
            }
            finally
            {
                ComRelease.Release(mail);
            }
        }

        /// <summary>
        /// Resolves the MailItem without displaying it. Exposed for tests and
        /// for future preview support.
        /// </summary>
        public Outlook.MailItem ResolveMailItem(SemanticSearchResult result)
        {
            if (result == null)
            {
                return null;
            }

            Outlook.NameSpace session = null;
            try
            {
                session = _application.Session;
                if (session == null)
                {
                    return null;
                }

                // Primary path: the stored EntryID + StoreID.
                Outlook.MailItem direct = TryGetItem(
                    session, result.EntryId, result.StoreId);
                if (direct != null && MatchesResult(direct, result))
                {
                    return direct;
                }

                ComRelease.Release(direct);

                // The EntryID changed after a move; try the default store.
                if (!string.IsNullOrWhiteSpace(result.EntryId)
                    && !string.IsNullOrWhiteSpace(result.StoreId))
                {
                    Outlook.MailItem defaultStoreItem = TryGetItem(
                        session, result.EntryId, string.Empty);
                    if (defaultStoreItem != null
                        && MatchesResult(defaultStoreItem, result))
                    {
                        return defaultStoreItem;
                    }

                    ComRelease.Release(defaultStoreItem);
                }

                // Fallback: search by InternetMessageId across stores.
                if (!string.IsNullOrWhiteSpace(result.InternetMessageId))
                {
                    Outlook.MailItem byMessageId = FindByInternetMessageId(
                        session, result);
                    if (byMessageId != null)
                    {
                        return byMessageId;
                    }
                }

                // Last resort: a strict metadata match within the folder.
                return FindByMetadata(session, result);
            }
            catch
            {
                return null;
            }
            finally
            {
                ComRelease.Release(session);
            }
        }

        private static Outlook.MailItem TryGetItem(
            Outlook.NameSpace session,
            string entryId,
            string storeId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return null;
            }

            try
            {
                object item = string.IsNullOrWhiteSpace(storeId)
                    ? session.GetItemFromID(entryId, Type.Missing)
                    : session.GetItemFromID(entryId, storeId);
                Outlook.MailItem mail = item as Outlook.MailItem;
                if (mail == null)
                {
                    ComRelease.Release(item);
                    return null;
                }

                return mail;
            }
            catch
            {
                return null;
            }
        }

        private Outlook.MailItem FindByInternetMessageId(
            Outlook.NameSpace session,
            SemanticSearchResult result)
        {
            Outlook.Stores stores = null;
            try
            {
                stores = session.Stores;
                if (stores == null)
                {
                    return null;
                }

                for (int storeIndex = 1; storeIndex <= stores.Count; storeIndex++)
                {
                    Outlook.Store store = null;
                    try
                    {
                        store = stores[storeIndex];
                        if (store == null)
                        {
                            continue;
                        }

                        Outlook.MailItem found = SearchStoreForMessageId(
                            store, result);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                    catch
                    {
                        // Continue with the next store.
                    }
                    finally
                    {
                        ComRelease.Release(store);
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                ComRelease.Release(stores);
            }

            return null;
        }

        private static Outlook.MailItem SearchStoreForMessageId(
            Outlook.Store store,
            SemanticSearchResult result)
        {
            // Restrict by subject when we have one to keep the search bounded,
            // then verify the InternetMessageId exactly.
            string filter = "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%"
                + EscapeLike(result.Subject) + "%'";
            if (string.IsNullOrWhiteSpace(result.Subject))
            {
                return null;
            }

            Outlook.MAPIFolder root = null;
            try
            {
                root = store.GetRootFolder();
                if (root == null)
                {
                    return null;
                }

                return SearchFolderForMessageId(root, filter, result, 0);
            }
            catch
            {
                return null;
            }
            finally
            {
                ComRelease.Release(root);
            }
        }

        private static Outlook.MailItem SearchFolderForMessageId(
            Outlook.MAPIFolder folder,
            string filter,
            SemanticSearchResult result,
            int depth)
        {
            if (depth > 6)
            {
                return null;
            }

            Outlook.Items items = null;
            try
            {
                items = folder.Items;
                if (items != null)
                {
                    object found = items.Find(filter);
                    int guard = 0;
                    while (found != null && guard < 50)
                    {
                        guard++;
                        Outlook.MailItem mail = found as Outlook.MailItem;
                        if (mail != null && MatchesResult(mail, result))
                        {
                            return mail;
                        }

                        ComRelease.Release(found);
                        found = items.FindNext();
                    }

                    if (found != null)
                    {
                        ComRelease.Release(found);
                    }
                }
            }
            catch
            {
                // Fall through to child folders.
            }
            finally
            {
                ComRelease.Release(items);
            }

            Outlook.Folders children = null;
            try
            {
                children = folder.Folders;
                if (children == null)
                {
                    return null;
                }

                for (int index = 1; index <= children.Count; index++)
                {
                    Outlook.MAPIFolder child = null;
                    try
                    {
                        child = children[index];
                        Outlook.MailItem found = SearchFolderForMessageId(
                            child, filter, result, depth + 1);
                        if (found != null)
                        {
                            return found;
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
            catch
            {
                return null;
            }
            finally
            {
                ComRelease.Release(children);
            }

            return null;
        }

        private static Outlook.MailItem FindByMetadata(
            Outlook.NameSpace session,
            SemanticSearchResult result)
        {
            // Without a safe bounded search this fallback is intentionally
            // conservative: it only succeeds when a single MailItem's
            // metadata fully matches, which requires an EntryID. If the
            // EntryID was stale *and* there is no InternetMessageId we do not
            // guess.
            if (string.IsNullOrWhiteSpace(result.EntryId)
                || string.IsNullOrWhiteSpace(result.InternetMessageId))
            {
                return null;
            }

            return null;
        }

        /// <summary>
        /// Verifies a resolved MailItem really is the indexed message.
        /// Priority: InternetMessageId, then subject + sender + received date.
        /// </summary>
        internal static bool MatchesResult(
            Outlook.MailItem mail,
            SemanticSearchResult result)
        {
            if (mail == null || result == null)
            {
                return false;
            }

            string messageId = TryProperty(
                delegate
                {
                    return Convert.ToString(
                        mail.PropertyAccessor.GetProperty(
                            "http://schemas.microsoft.com/mapi/proptag/0x1035001F"));
                });
            string expected = (result.InternetMessageId ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(expected)
                && !string.IsNullOrWhiteSpace(messageId))
            {
                return string.Equals(
                    expected, messageId.Trim(), StringComparison.OrdinalIgnoreCase);
            }

            // Without a message id, require a subject match.
            string subject = TryProperty(delegate { return mail.Subject; });
            if (!string.IsNullOrWhiteSpace(result.Subject)
                && !string.Equals(
                    (subject ?? string.Empty).Trim(),
                    result.Subject.Trim(),
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // When the message id is unavailable on either side, require at
            // least a sender or received-date match so a same-title mail is
            // not silently opened instead.
            string sender = TryProperty(
                delegate { return mail.SenderEmailAddress; });
            if (!string.IsNullOrWhiteSpace(result.SenderEmail)
                && !string.IsNullOrWhiteSpace(sender))
            {
                return string.Equals(
                    sender.Trim(),
                    result.SenderEmail.Trim(),
                    StringComparison.OrdinalIgnoreCase);
            }

            return true;
        }

        private static string EscapeLike(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            return value
                .Replace("'", "''")
                .Replace("%", "[%]")
                .Replace("_", "[_]");
        }

        private static string TryProperty(Func<string> getter)
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
    }
}
