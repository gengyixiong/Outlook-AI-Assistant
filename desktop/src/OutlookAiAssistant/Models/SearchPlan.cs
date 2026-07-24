using System.Collections.Generic;

namespace OutlookAiAssistant.Models
{
    /// <summary>
    /// Allow-listed intermediate representation produced by the language model.
    /// Only these fields can be translated into an Outlook local search query.
    /// </summary>
    public sealed class SearchPlan
    {
        public List<string> From { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public bool ToMe { get; set; }
        public List<List<string>> TextGroups { get; set; }
        public List<List<string>> SubjectGroups { get; set; }
        public List<List<string>> BodyGroups { get; set; }
        public string ReceivedFrom { get; set; }
        public string ReceivedThrough { get; set; }
        public bool? HasAttachments { get; set; }
        public bool? IsUnread { get; set; }
        public string Scope { get; set; }

        public SearchPlan()
        {
            From = new List<string>();
            To = new List<string>();
            Cc = new List<string>();
            TextGroups = new List<List<string>>();
            SubjectGroups = new List<List<string>>();
            BodyGroups = new List<List<string>>();
            ReceivedFrom = string.Empty;
            ReceivedThrough = string.Empty;
            Scope = "all_folders";
        }

        public void Normalize()
        {
            From = From ?? new List<string>();
            To = To ?? new List<string>();
            Cc = Cc ?? new List<string>();
            TextGroups = TextGroups ?? new List<List<string>>();
            SubjectGroups = SubjectGroups ?? new List<List<string>>();
            BodyGroups = BodyGroups ?? new List<List<string>>();
            ReceivedFrom = ReceivedFrom ?? string.Empty;
            ReceivedThrough = ReceivedThrough ?? string.Empty;
            Scope = string.IsNullOrWhiteSpace(Scope) ? "all_folders" : Scope.Trim();
        }
    }
}

