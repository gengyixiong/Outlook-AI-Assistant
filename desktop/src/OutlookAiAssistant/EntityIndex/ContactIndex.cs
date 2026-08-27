using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAiAssistant.EntityIndex
{
    public sealed class ContactIndex
    {
        public int Version { get; set; }
        public string CreatedAt { get; set; }
        public List<ContactProfile> Contacts { get; set; }
        public List<FolderContextEntry> FolderContext { get; set; }

        public ContactIndex()
        {
            Version = 1;
            CreatedAt = string.Empty;
            Contacts = new List<ContactProfile>();
            FolderContext = new List<FolderContextEntry>();
        }
    }

    public sealed class ContactProfile
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string Company { get; set; }
        public string Country { get; set; }
        public string City { get; set; }
        public string Role { get; set; }
        public string JobTitle { get; set; }
        public List<string> Aliases { get; set; }

        public ContactProfile()
        {
            Id = string.Empty;
            DisplayName = string.Empty;
            Email = string.Empty;
            Company = string.Empty;
            Country = string.Empty;
            City = string.Empty;
            Role = string.Empty;
            JobTitle = string.Empty;
            Aliases = new List<string>();
        }

        public static string CreateId(string email)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(
                (email ?? string.Empty).Trim().ToLowerInvariant());
            byte[] hash;
            using (SHA256 algorithm = SHA256.Create())
            {
                hash = algorithm.ComputeHash(bytes);
            }

            StringBuilder id = new StringBuilder("contact_");
            for (int index = 0; index < 8; index++)
            {
                id.Append(hash[index].ToString("x2"));
            }

            return id.ToString();
        }
    }

    public sealed class FolderContextEntry
    {
        public string Root { get; set; }
        public string RelativePath { get; set; }
        public List<string> ContactIds { get; set; }

        public FolderContextEntry()
        {
            Root = string.Empty;
            RelativePath = string.Empty;
            ContactIds = new List<string>();
        }
    }
}
