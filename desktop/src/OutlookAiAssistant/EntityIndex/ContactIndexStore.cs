using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Web.Script.Serialization;

namespace OutlookAiAssistant.EntityIndex
{
    /// <summary>
    /// Owns the single formal contact index and replaces it only after the new
    /// JSON has been serialized, written and parsed successfully.
    /// </summary>
    public sealed class ContactIndexStore
    {
        private const int MaximumIndexLength = 4 * 1024 * 1024;
        private readonly JavaScriptSerializer _serializer;

        public string DirectoryPath { get; private set; }
        public string IndexPath { get; private set; }
        public string BackupDirectory { get; private set; }

        public ContactIndexStore()
            : this(Path.Combine(
                Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData),
                "OutlookAiAssistant"))
        {
        }

        public ContactIndexStore(string directoryPath)
        {
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                throw new ArgumentException(
                    "Contact index directory is required.",
                    "directoryPath");
            }

            DirectoryPath = Path.GetFullPath(directoryPath);
            IndexPath = Path.Combine(DirectoryPath, "contact-index.json");
            BackupDirectory = Path.Combine(DirectoryPath, "backup");
            _serializer = new JavaScriptSerializer();
            _serializer.MaxJsonLength = 4 * 1024 * 1024;
        }

        public bool Exists()
        {
            return File.Exists(IndexPath);
        }

        public ContactIndex Load()
        {
            if (!File.Exists(IndexPath))
            {
                return new ContactIndex();
            }

            if (new FileInfo(IndexPath).Length > MaximumIndexLength)
            {
                throw new InvalidOperationException(
                    "联系人索引超过安全大小限制。");
            }

            return Parse(File.ReadAllText(IndexPath, Encoding.UTF8));
        }

        public string Serialize(ContactIndex index)
        {
            ValidateAndNormalize(index);
            Dictionary<string, object> root = new Dictionary<string, object>();
            root["version"] = index.Version;
            root["created_at"] = index.CreatedAt;

            List<object> contacts = new List<object>();
            for (int contactIndex = 0;
                contactIndex < index.Contacts.Count;
                contactIndex++)
            {
                ContactProfile contact = index.Contacts[contactIndex];
                Dictionary<string, object> item =
                    new Dictionary<string, object>();
                item["id"] = contact.Id;
                item["display_name"] = contact.DisplayName;
                item["email"] = contact.Email;
                item["company"] = contact.Company;
                item["country"] = contact.Country;
                item["city"] = contact.City;
                item["role"] = contact.Role;
                item["job_title"] = contact.JobTitle;
                item["aliases"] = contact.Aliases.ToArray();
                contacts.Add(item);
            }

            List<object> folderContext = new List<object>();
            for (int folderIndex = 0;
                folderIndex < index.FolderContext.Count;
                folderIndex++)
            {
                FolderContextEntry entry = index.FolderContext[folderIndex];
                Dictionary<string, object> item =
                    new Dictionary<string, object>();
                item["root"] = entry.Root;
                item["relative_path"] = entry.RelativePath;
                item["contact_ids"] = entry.ContactIds.ToArray();
                folderContext.Add(item);
            }

            root["contacts"] = contacts.ToArray();
            root["folder_context"] = folderContext.ToArray();
            return _serializer.Serialize(root);
        }

        public ContactIndex Parse(string json)
        {
            if (json == null || json.Length > MaximumIndexLength)
            {
                throw new InvalidOperationException(
                    "联系人索引超过安全大小限制。");
            }

            Dictionary<string, object> root;
            try
            {
                root = _serializer.DeserializeObject(json ?? string.Empty)
                    as Dictionary<string, object>;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    "联系人索引不是有效的 JSON。",
                    ex);
            }

            if (root == null)
            {
                throw new InvalidOperationException("联系人索引内容为空。");
            }

            ContactIndex index = new ContactIndex();
            index.Version = ReadInt(root, "version", 0);
            index.CreatedAt = ReadString(root, "created_at");

            object rawContacts;
            object[] contacts = root.TryGetValue("contacts", out rawContacts)
                ? rawContacts as object[]
                : null;
            if (contacts == null)
            {
                throw new InvalidOperationException(
                    "联系人索引缺少 contacts 数组。");
            }

            for (int itemIndex = 0;
                itemIndex < contacts.Length;
                itemIndex++)
            {
                Dictionary<string, object> item = contacts[itemIndex]
                    as Dictionary<string, object>;
                if (item == null)
                {
                    throw new InvalidOperationException(
                        "contacts 数组包含无效条目。");
                }

                ContactProfile contact = new ContactProfile();
                contact.Id = ReadString(item, "id");
                contact.DisplayName = ReadString(item, "display_name");
                contact.Email = ReadString(item, "email");
                contact.Company = ReadString(item, "company");
                contact.Country = ReadString(item, "country");
                contact.City = ReadString(item, "city");
                contact.Role = ReadString(item, "role");
                contact.JobTitle = ReadString(item, "job_title");
                contact.Aliases = ReadStringList(item, "aliases");
                index.Contacts.Add(contact);
            }

            object rawFolderContext;
            object[] folderContext = root.TryGetValue(
                "folder_context",
                out rawFolderContext)
                    ? rawFolderContext as object[]
                    : null;
            if (root.ContainsKey("folder_context")
                && rawFolderContext != null
                && folderContext == null)
            {
                throw new InvalidOperationException(
                    "folder_context 必须是数组。");
            }

            if (folderContext != null)
            {
                for (int itemIndex = 0;
                    itemIndex < folderContext.Length;
                    itemIndex++)
                {
                    Dictionary<string, object> item = folderContext[itemIndex]
                        as Dictionary<string, object>;
                    if (item == null)
                    {
                        throw new InvalidOperationException(
                            "folder_context 数组包含无效条目。");
                    }

                    FolderContextEntry entry = new FolderContextEntry();
                    entry.Root = ReadString(item, "root");
                    entry.RelativePath = ReadString(item, "relative_path");
                    entry.ContactIds = ReadStringList(item, "contact_ids");
                    index.FolderContext.Add(entry);
                }
            }

            ValidateAndNormalize(index);
            return index;
        }

        public void Save(ContactIndex index)
        {
            string json = Serialize(index);
            Directory.CreateDirectory(DirectoryPath);
            string temporaryPath = Path.Combine(
                DirectoryPath,
                "contact-index.new.json");

            try
            {
                File.WriteAllText(
                    temporaryPath,
                    json,
                    new UTF8Encoding(false));
                Parse(File.ReadAllText(temporaryPath, Encoding.UTF8));

                if (File.Exists(IndexPath))
                {
                    Directory.CreateDirectory(BackupDirectory);
                    string backupPath = CreateBackupPath();
                    File.Copy(IndexPath, backupPath, false);
                    File.Replace(temporaryPath, IndexPath, null);
                }
                else
                {
                    File.Move(temporaryPath, IndexPath);
                }
            }
            finally
            {
                if (File.Exists(temporaryPath))
                {
                    File.Delete(temporaryPath);
                }
            }
        }

        private string CreateBackupPath()
        {
            string stamp = DateTime.Now.ToString(
                "yyyyMMdd-HHmmss",
                CultureInfo.InvariantCulture);
            string path = Path.Combine(
                BackupDirectory,
                "contact-index-" + stamp + ".json");
            int suffix = 1;
            while (File.Exists(path))
            {
                path = Path.Combine(
                    BackupDirectory,
                    "contact-index-" + stamp + "-" + suffix + ".json");
                suffix++;
            }

            return path;
        }

        private static void ValidateAndNormalize(ContactIndex index)
        {
            if (index == null)
            {
                throw new ArgumentNullException("index");
            }

            if (index.Version != 1)
            {
                throw new InvalidOperationException(
                    "不支持的联系人索引版本。");
            }

            DateTime createdAt;
            if (!DateTime.TryParse(
                index.CreatedAt,
                CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind,
                out createdAt))
            {
                throw new InvalidOperationException(
                    "联系人索引缺少有效的创建时间。");
            }

            index.Contacts = index.Contacts ?? new List<ContactProfile>();
            index.FolderContext = index.FolderContext
                ?? new List<FolderContextEntry>();
            if (index.Contacts.Count > 1000
                || index.FolderContext.Count > 5000)
            {
                throw new InvalidOperationException("联系人索引超过安全大小限制。");
            }

            HashSet<string> ids = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            HashSet<string> emails = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (int contactIndex = index.Contacts.Count - 1;
                contactIndex >= 0;
                contactIndex--)
            {
                ContactProfile contact = index.Contacts[contactIndex];
                if (contact == null)
                {
                    index.Contacts.RemoveAt(contactIndex);
                    continue;
                }

                contact.Id = Trim(contact.Id, 80);
                contact.DisplayName = Trim(contact.DisplayName, 200);
                contact.Email = Trim(contact.Email, 320).ToLowerInvariant();
                contact.Company = Trim(contact.Company, 200);
                contact.Country = Trim(contact.Country, 100);
                contact.City = Trim(contact.City, 100);
                contact.Role = Trim(contact.Role, 100);
                contact.JobTitle = Trim(contact.JobTitle, 200);
                contact.Aliases = contact.Aliases ?? new List<string>();
                NormalizeAliases(contact);

                if (!IsContactId(contact.Id)
                    || !IsEmail(contact.Email)
                    || !string.Equals(
                        contact.Id,
                        ContactProfile.CreateId(contact.Email),
                        StringComparison.Ordinal)
                    || !ids.Add(contact.Id)
                    || !emails.Add(contact.Email))
                {
                    throw new InvalidOperationException(
                        "联系人索引包含无效或重复的联系人标识。");
                }
            }

            for (int folderIndex = index.FolderContext.Count - 1;
                folderIndex >= 0;
                folderIndex--)
            {
                FolderContextEntry entry = index.FolderContext[folderIndex];
                if (entry == null)
                {
                    index.FolderContext.RemoveAt(folderIndex);
                    continue;
                }

                entry.Root = Trim(entry.Root, 260);
                entry.RelativePath = Trim(entry.RelativePath, 1000);
                entry.ContactIds = entry.ContactIds ?? new List<string>();
                NormalizeContactIds(entry.ContactIds, ids);
                if (entry.Root.Length == 0 || entry.RelativePath.Length == 0)
                {
                    index.FolderContext.RemoveAt(folderIndex);
                }
            }
        }

        private static void NormalizeContactIds(
            List<string> contactIds,
            ISet<string> validContactIds)
        {
            HashSet<string> seen = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            for (int index = contactIds.Count - 1; index >= 0; index--)
            {
                string id = Trim(contactIds[index], 80);
                if (!validContactIds.Contains(id) || !seen.Add(id))
                {
                    contactIds.RemoveAt(index);
                }
                else
                {
                    contactIds[index] = id;
                }
            }

            contactIds.Sort(StringComparer.OrdinalIgnoreCase);
        }

        private static void NormalizeAliases(ContactProfile contact)
        {
            List<string> normalized = new List<string>();
            HashSet<string> seen = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            AddAlias(normalized, seen, contact.DisplayName);
            AddAlias(normalized, seen, contact.Email);
            int at = contact.Email.IndexOf('@');
            if (at > 0)
            {
                AddAlias(normalized, seen, contact.Email.Substring(0, at));
            }

            for (int aliasIndex = 0;
                aliasIndex < contact.Aliases.Count && normalized.Count < 50;
                aliasIndex++)
            {
                AddAlias(normalized, seen, contact.Aliases[aliasIndex]);
            }

            contact.Aliases = normalized;
        }

        private static void AddAlias(
            ICollection<string> aliases,
            ISet<string> seen,
            string value)
        {
            string alias = Trim(value, 200);
            if (alias.Length > 0 && seen.Add(alias))
            {
                aliases.Add(alias);
            }
        }

        private static bool IsContactId(string value)
        {
            if (string.IsNullOrWhiteSpace(value)
                || !value.StartsWith("contact_", StringComparison.Ordinal))
            {
                return false;
            }

            for (int index = "contact_".Length; index < value.Length; index++)
            {
                char character = value[index];
                if (!char.IsLetterOrDigit(character)
                    && character != '_'
                    && character != '-')
                {
                    return false;
                }
            }

            return value.Length > "contact_".Length;
        }

        private static bool IsEmail(string value)
        {
            try
            {
                MailAddress address = new MailAddress(value);
                return string.Equals(
                    address.Address,
                    value,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static string Trim(string value, int maximumLength)
        {
            string result = (value ?? string.Empty).Trim();
            return result.Length <= maximumLength
                ? result
                : result.Substring(0, maximumLength);
        }

        private static string ReadString(
            IDictionary<string, object> source,
            string key)
        {
            object value;
            return source.TryGetValue(key, out value) && value != null
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private static int ReadInt(
            IDictionary<string, object> source,
            string key,
            int defaultValue)
        {
            object value;
            int result;
            return source.TryGetValue(key, out value)
                && value != null
                && int.TryParse(
                    Convert.ToString(value, CultureInfo.InvariantCulture),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out result)
                        ? result
                        : defaultValue;
        }

        private static List<string> ReadStringList(
            IDictionary<string, object> source,
            string key)
        {
            List<string> values = new List<string>();
            object raw;
            object[] items = source.TryGetValue(key, out raw)
                ? raw as object[]
                : null;
            if (items == null && raw != null)
            {
                throw new InvalidOperationException(
                    key + " 必须是字符串数组。");
            }

            if (items == null)
            {
                return values;
            }

            for (int index = 0; index < items.Length; index++)
            {
                if (items[index] == null)
                {
                    continue;
                }

                string value = items[index] as string;
                if (value == null)
                {
                    throw new InvalidOperationException(
                        key + " 必须只包含字符串。");
                }

                values.Add(value);
            }

            return values;
        }
    }
}
