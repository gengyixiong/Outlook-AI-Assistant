using System.Collections.Generic;

namespace OutlookAiAssistant.Models
{
    /// <summary>
    /// Allow-listed intermediate representation produced by the language model.
    /// Only these fields can be translated into an Outlook local search query.
    /// </summary>
    public sealed class SearchPlan
    {
        public List<string> MatchedContactIds { get; set; }
        public List<string> From { get; set; }
        public List<string> To { get; set; }
        public List<string> Cc { get; set; }
        public bool ToMe { get; set; }
        /// <summary>
        /// Rare, literal identifiers that are likely to occur in the target
        /// message, for example a product name, project code or email address.
        /// Each inner list contains spelling aliases joined with OR.
        /// </summary>
        public List<List<string>> AnchorGroups { get; set; }

        /// <summary>
        /// Business concepts and their multilingual synonyms. These are added
        /// only after the broad anchor-only query.
        /// </summary>
        public List<List<string>> ConceptGroups { get; set; }

        /// <summary>
        /// Context such as country, customer type or relationship. Hints are
        /// intentionally reserved for the precise query because they may not
        /// literally occur in the message.
        /// </summary>
        public List<List<string>> HintGroups { get; set; }

        // Legacy field retained so older compatible model responses can still
        // be parsed. New prompts use ConceptGroups instead.
        public List<List<string>> TextGroups { get; set; }
        public List<List<string>> SubjectGroups { get; set; }
        public List<List<string>> BodyGroups { get; set; }
        public List<List<string>> AttachmentNameGroups { get; set; }
        public List<string> AttachmentExtensions { get; set; }
        public string ReceivedFrom { get; set; }
        public string ReceivedThrough { get; set; }
        public bool? HasAttachments { get; set; }
        public bool? IsUnread { get; set; }
        public string Scope { get; set; }

        public SearchPlan()
        {
            MatchedContactIds = new List<string>();
            From = new List<string>();
            To = new List<string>();
            Cc = new List<string>();
            AnchorGroups = new List<List<string>>();
            ConceptGroups = new List<List<string>>();
            HintGroups = new List<List<string>>();
            TextGroups = new List<List<string>>();
            SubjectGroups = new List<List<string>>();
            BodyGroups = new List<List<string>>();
            AttachmentNameGroups = new List<List<string>>();
            AttachmentExtensions = new List<string>();
            ReceivedFrom = string.Empty;
            ReceivedThrough = string.Empty;
            Scope = "all_folders";
        }

        public void Normalize()
        {
            MatchedContactIds = MatchedContactIds ?? new List<string>();
            From = From ?? new List<string>();
            To = To ?? new List<string>();
            Cc = Cc ?? new List<string>();
            AnchorGroups = AnchorGroups ?? new List<List<string>>();
            ConceptGroups = ConceptGroups ?? new List<List<string>>();
            HintGroups = HintGroups ?? new List<List<string>>();
            TextGroups = TextGroups ?? new List<List<string>>();
            SubjectGroups = SubjectGroups ?? new List<List<string>>();
            BodyGroups = BodyGroups ?? new List<List<string>>();
            AttachmentNameGroups = AttachmentNameGroups
                ?? new List<List<string>>();
            AttachmentExtensions = AttachmentExtensions ?? new List<string>();
            ReceivedFrom = ReceivedFrom ?? string.Empty;
            ReceivedThrough = ReceivedThrough ?? string.Empty;
            Scope = string.IsNullOrWhiteSpace(Scope) ? "all_folders" : Scope.Trim();
        }
    }
}

