using System;
using System.Collections.Generic;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Plain-data payload for one mail sent to the local embedding service.
    /// No COM object is referenced after the snapshot is created.
    /// </summary>
    public sealed class MailIndexPayload
    {
        public string StoreId { get; set; }
        public string EntryId { get; set; }
        public string Subject { get; set; }
        public string InternetMessageId { get; set; }
        public string ConversationId { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string Recipients { get; set; }
        public string ReceivedAt { get; set; }
        public string SentAt { get; set; }
        public string LastModifiedAt { get; set; }
        public string FolderId { get; set; }
        public string FolderPath { get; set; }
        public string AttachmentNames { get; set; }
        public int AttachmentCount { get; set; }
        public string BodyText { get; set; }
        public string HtmlBody { get; set; }
        public string OwnerMailbox { get; set; }
        public bool ForceReembed { get; set; }

        public MailIndexPayload()
        {
            StoreId = string.Empty;
            EntryId = string.Empty;
            Subject = string.Empty;
            InternetMessageId = string.Empty;
            ConversationId = string.Empty;
            SenderName = string.Empty;
            SenderEmail = string.Empty;
            Recipients = string.Empty;
            ReceivedAt = string.Empty;
            SentAt = string.Empty;
            LastModifiedAt = string.Empty;
            FolderId = string.Empty;
            FolderPath = string.Empty;
            AttachmentNames = string.Empty;
            BodyText = string.Empty;
            HtmlBody = string.Empty;
            OwnerMailbox = string.Empty;
        }
    }

    /// <summary>One search result returned by the local service.</summary>
    public sealed class SemanticSearchResult
    {
        public int MailId { get; set; }
        public string StoreId { get; set; }
        public string EntryId { get; set; }
        public string InternetMessageId { get; set; }
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public string SenderEmail { get; set; }
        public string ReceivedAt { get; set; }
        public string FolderPath { get; set; }
        public string AttachmentNames { get; set; }
        public int AttachmentCount { get; set; }
        public string Snippet { get; set; }
        public double Score { get; set; }
        public int? SemanticRank { get; set; }
        public int? KeywordRank { get; set; }

        public SemanticSearchResult()
        {
            StoreId = string.Empty;
            EntryId = string.Empty;
            InternetMessageId = string.Empty;
            Subject = string.Empty;
            SenderName = string.Empty;
            SenderEmail = string.Empty;
            ReceivedAt = string.Empty;
            FolderPath = string.Empty;
            AttachmentNames = string.Empty;
            Snippet = string.Empty;
        }
    }

    /// <summary>Response envelope for a hybrid search request.</summary>
    public sealed class SemanticSearchResponse
    {
        public string Query { get; set; }
        public IList<SemanticSearchResult> Results { get; set; }
        public bool SemanticAvailable { get; set; }
        public bool KeywordAvailable { get; set; }
        public int CandidateCount { get; set; }
        public bool FilterApplied { get; set; }

        public SemanticSearchResponse()
        {
            Query = string.Empty;
            Results = new List<SemanticSearchResult>();
        }
    }

    /// <summary>Health and index statistics from the local service.</summary>
    public sealed class SemanticServiceStatus
    {
        public bool Reachable { get; set; }
        public string Status { get; set; }
        public string ModelId { get; set; }
        public bool ModelAvailable { get; set; }
        /// <summary>True when the model files are already cached locally.</summary>
        public bool ModelCached { get; set; }
        public string ModelError { get; set; }
        public string Device { get; set; }
        /// <summary>Why this device was chosen (auto-detection reason).</summary>
        public string DeviceReason { get; set; }
        public int MailCount { get; set; }
        public int ChunkCount { get; set; }
        public int EmbeddedChunkCount { get; set; }
        public long SizeBytes { get; set; }
        public string LastSyncAt { get; set; }
        public string Message { get; set; }

        // Local path hints reported by the service, used by the settings UI.
        public string ModelDir { get; set; }
        public string IndexDir { get; set; }
        public string DataDir { get; set; }
        public string LogDir { get; set; }
        public string ServiceDir { get; set; }

        public SemanticServiceStatus()
        {
            Status = string.Empty;
            ModelId = string.Empty;
            ModelError = string.Empty;
            Device = string.Empty;
            DeviceReason = string.Empty;
            LastSyncAt = string.Empty;
            Message = string.Empty;
            ModelDir = string.Empty;
            IndexDir = string.Empty;
            DataDir = string.Empty;
            LogDir = string.Empty;
            ServiceDir = string.Empty;
        }
    }

    /// <summary>Ranked result paired with the Outlook item locator outcome.</summary>
    public sealed class LocatedSearchResult
    {
        public SemanticSearchResult Result { get; set; }
        public bool ItemResolved { get; set; }
        public string LocationMessage { get; set; }

        public LocatedSearchResult()
        {
            LocationMessage = string.Empty;
        }
    }
}
