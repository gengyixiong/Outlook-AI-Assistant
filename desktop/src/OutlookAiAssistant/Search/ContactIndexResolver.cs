using System;
using System.Collections.Generic;
using OutlookAiAssistant.EntityIndex;
using OutlookAiAssistant.Models;

namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// Resolves only exact IDs from the validated local index. Folder context
    /// and model-generated IDs can never become an Outlook identity directly.
    /// </summary>
    public sealed class ContactIndexResolver
    {
        public void Resolve(SearchPlan plan, ContactIndex index)
        {
            if (plan == null)
            {
                throw new ArgumentNullException("plan");
            }

            plan.Normalize();
            Dictionary<string, ContactProfile> byId =
                new Dictionary<string, ContactProfile>(
                    StringComparer.OrdinalIgnoreCase);
            if (index != null && index.Contacts != null)
            {
                for (int contactIndex = 0;
                    contactIndex < index.Contacts.Count;
                    contactIndex++)
                {
                    ContactProfile contact = index.Contacts[contactIndex];
                    if (contact != null
                        && !string.IsNullOrWhiteSpace(contact.Id)
                        && !string.IsNullOrWhiteSpace(contact.Email))
                    {
                        byId[contact.Id] = contact;
                    }
                }
            }

            List<string> validIds = new List<string>();
            HashSet<string> seenIds = new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);
            HashSet<string> emails = new HashSet<string>(
                plan.From,
                StringComparer.OrdinalIgnoreCase);
            for (int idIndex = 0;
                idIndex < plan.MatchedContactIds.Count;
                idIndex++)
            {
                string id = plan.MatchedContactIds[idIndex];
                ContactProfile contact;
                if (!byId.TryGetValue(id, out contact))
                {
                    continue;
                }

                if (seenIds.Add(contact.Id))
                {
                    validIds.Add(contact.Id);
                }
                if (emails.Add(contact.Email))
                {
                    plan.From.Add(contact.Email);
                }
            }

            plan.MatchedContactIds = validIds;
        }
    }
}
