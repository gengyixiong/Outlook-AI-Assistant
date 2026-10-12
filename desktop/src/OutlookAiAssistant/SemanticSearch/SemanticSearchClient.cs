using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// Thrown when the local semantic service cannot be reached or returns an
    /// error. The message is user-friendly; the add-in must not crash.
    /// </summary>
    public sealed class SemanticServiceException : Exception
    {
        /// <summary>HTTP status code, or 0 when the request never connected.</summary>
        public int StatusCode { get; private set; }

        public bool IsModelUnavailable
        {
            get { return StatusCode == 503; }
        }

        public SemanticServiceException(string message)
            : base(message)
        {
        }

        public SemanticServiceException(string message, Exception inner)
            : base(message, inner)
        {
        }

        public SemanticServiceException(string message, int statusCode)
            : base(message)
        {
            StatusCode = statusCode;
        }
    }

    /// <summary>
    /// HTTP client for the loopback-only Python semantic search service.
    ///
    /// The service loads EmbeddingGemma 2 and owns the SQLite index. This
    /// client never reads the database directly and never sends mail content
    /// anywhere except 127.0.0.1.
    /// </summary>
    public sealed class SemanticSearchClient
    {
        private const string DefaultBaseUrl = "http://127.0.0.1:8765";
        private const int RequestTimeoutMilliseconds = 120000;

        private readonly JavaScriptSerializer _serializer;
        private readonly string _baseUrl;
        private readonly string _token;

        public SemanticSearchClient(string baseUrl, string token)
        {
            _serializer = new JavaScriptSerializer();
            _serializer.MaxJsonLength = 32 * 1024 * 1024;
            _baseUrl = NormalizeBaseUrl(baseUrl);
            _token = token ?? string.Empty;
        }

        public string BaseUrl
        {
            get { return _baseUrl; }
        }

        /// <summary>Reads the shared token the service writes on first start.</summary>
        public static string ReadServiceToken()
        {
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "OutlookAiAssistant",
                    "semantic-service-token");
                if (!File.Exists(path))
                {
                    return string.Empty;
                }

                return File.ReadAllText(path, Encoding.UTF8).Trim();
            }
            catch
            {
                return string.Empty;
            }
        }

        public Task<SemanticServiceStatus> GetStatusAsync(
            CancellationToken cancellationToken)
        {
            return Task.Factory.StartNew(
                delegate { return GetStatus(); },
                cancellationToken);
        }

        /// <summary>
        /// Downloads and caches the embedding model on the local machine. This
        /// is the explicit step a fresh machine needs; without it every index
        /// request fails with a clear model-unavailable error.
        /// </summary>
        public void DownloadModel()
        {
            Send("POST", "/model/download", "{}", true);
        }

        /// <summary>Reports whether the model can be used, and why not.</summary>
        public SemanticServiceStatus GetModelStatus()
        {
            try
            {
                string json = Send("GET", "/model/status", null, true);
                Dictionary<string, object> root = Deserialize(json);
                SemanticServiceStatus status = new SemanticServiceStatus();
                status.Reachable = true;
                status.ModelId = GetString(root, "model_id");
                status.ModelAvailable = GetBool(root, "available");
                status.ModelError = GetString(root, "error");
                return status;
            }
            catch (Exception ex)
            {
                SemanticServiceStatus status = new SemanticServiceStatus();
                status.Reachable = false;
                status.Message = ex.Message;
                return status;
            }
        }

        public SemanticServiceStatus GetStatus()
        {
            try
            {
                string json = Send("GET", "/health", null, false);
                Dictionary<string, object> root = Deserialize(json);
                Dictionary<string, object> index =
                    GetObject(root, "index");
                SemanticServiceStatus status = new SemanticServiceStatus();
                status.Reachable = true;
                status.Status = GetString(root, "status");
                status.ModelId = GetString(root, "model_id");
                status.ModelAvailable = GetBool(root, "model_available");
                status.ModelCached = GetBool(root, "model_cached");
                status.ModelError = GetString(root, "model_error");
                status.Device = GetString(root, "device");
                status.DeviceReason = GetString(root, "device_reason");
                if (index != null)
                {
                    status.MailCount = GetInt(index, "mail_count");
                    status.ChunkCount = GetInt(index, "chunk_count");
                    status.EmbeddedChunkCount =
                        GetInt(index, "embedded_chunk_count");
                    status.SizeBytes = GetLong(index, "size_bytes");
                    status.LastSyncAt = GetString(index, "last_sync_at");
                }

                Dictionary<string, object> paths =
                    GetObject(root, "paths");
                if (paths != null)
                {
                    status.ModelDir = GetString(paths, "model_dir");
                    status.IndexDir = GetString(paths, "index_dir");
                    status.DataDir = GetString(paths, "data_dir");
                    status.LogDir = GetString(paths, "log_dir");
                    status.ServiceDir = GetString(paths, "service_dir");
                }

                return status;
            }
            catch (Exception ex)
            {
                SemanticServiceStatus status = new SemanticServiceStatus();
                status.Reachable = false;
                status.Message = ex.Message;
                return status;
            }
        }

        public Task<string> IndexMailAsync(
            MailIndexPayload payload,
            CancellationToken cancellationToken)
        {
            if (payload == null)
            {
                throw new ArgumentNullException("payload");
            }

            return Task.Factory.StartNew(
                delegate
                {
                    string body = _serializer.Serialize(
                        ToSnakeCaseDictionary(payload));
                    return Send("POST", "/index/mail", body, true);
                },
                cancellationToken);
        }

        /// <summary>
        /// JavaScriptSerializer preserves CLR property names, but the service
        /// expects snake_case. Build the payload explicitly so the two sides
        /// cannot silently disagree on field names.
        /// </summary>
        public static Dictionary<string, object> ToSnakeCaseDictionary(
            MailIndexPayload payload)
        {
            Dictionary<string, object> map =
                new Dictionary<string, object>();
            map["store_id"] = payload.StoreId ?? string.Empty;
            map["entry_id"] = payload.EntryId ?? string.Empty;
            map["subject"] = payload.Subject ?? string.Empty;
            map["internet_message_id"] =
                payload.InternetMessageId ?? string.Empty;
            map["conversation_id"] = payload.ConversationId ?? string.Empty;
            map["sender_name"] = payload.SenderName ?? string.Empty;
            map["sender_email"] = payload.SenderEmail ?? string.Empty;
            map["recipients"] = payload.Recipients ?? string.Empty;
            map["received_at"] = payload.ReceivedAt ?? string.Empty;
            map["sent_at"] = payload.SentAt ?? string.Empty;
            map["last_modified_at"] = payload.LastModifiedAt ?? string.Empty;
            map["folder_id"] = payload.FolderId ?? string.Empty;
            map["folder_path"] = payload.FolderPath ?? string.Empty;
            map["attachment_names"] = payload.AttachmentNames ?? string.Empty;
            map["attachment_count"] = payload.AttachmentCount;
            map["body_text"] = payload.BodyText ?? string.Empty;
            map["html_body"] = payload.HtmlBody ?? string.Empty;
            map["owner_mailbox"] = payload.OwnerMailbox ?? string.Empty;
            map["force_reembed"] = payload.ForceReembed;
            return map;
        }

        public void DeleteMail(string storeId, string entryId)
        {
            Dictionary<string, object> payload =
                new Dictionary<string, object>();
            payload["store_id"] = storeId ?? string.Empty;
            payload["entry_id"] = entryId ?? string.Empty;
            Send("POST", "/index/delete", _serializer.Serialize(payload), true);
        }

        /// <summary>
        /// Asks the service to stop itself. The response arrives before the
        /// process exits, so the caller should not treat a missing reply as a
        /// failure.
        /// </summary>
        public void ShutdownService()
        {
            Send("POST", "/service/shutdown", "{}", true);
        }

        public void ClearIndex()
        {
            Send("DELETE", "/index", null, true);
        }

        public void ReportSync(string modelId)
        {
            try
            {
                Dictionary<string, object> payload =
                    new Dictionary<string, object>();
                payload["last_sync_at"] = DateTime.UtcNow.ToString("o");
                if (!string.IsNullOrWhiteSpace(modelId))
                {
                    payload["model_id"] = modelId;
                }

                Send(
                    "POST",
                    "/index/meta",
                    _serializer.Serialize(payload),
                    true);
            }
            catch
            {
                // Sync metadata is advisory; failure must not fail indexing.
            }
        }

        public Task<SemanticSearchResponse> SearchAsync(
            string query,
            IList<string> senders,
            string receivedFrom,
            string receivedThrough,
            bool? hasAttachments,
            int? topK,
            CancellationToken cancellationToken)
        {
            return Task.Factory.StartNew(
                delegate
                {
                    return Search(
                        query,
                        senders,
                        receivedFrom,
                        receivedThrough,
                        hasAttachments,
                        topK);
                },
                cancellationToken);
        }

        public SemanticSearchResponse Search(
            string query,
            IList<string> senders,
            string receivedFrom,
            string receivedThrough,
            bool? hasAttachments,
            int? topK)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                throw new ArgumentException("搜索描述不能为空。", "query");
            }

            Dictionary<string, object> payload =
                new Dictionary<string, object>();
            payload["query"] = query;
            if (topK.HasValue)
            {
                payload["top_k"] = topK.Value;
            }

            if (senders != null && senders.Count > 0)
            {
                List<string> clean = new List<string>();
                for (int index = 0; index < senders.Count; index++)
                {
                    if (!string.IsNullOrWhiteSpace(senders[index]))
                    {
                        clean.Add(senders[index].Trim());
                    }
                }

                if (clean.Count > 0)
                {
                    payload["senders"] = clean;
                }
            }

            if (!string.IsNullOrWhiteSpace(receivedFrom))
            {
                payload["received_from"] = receivedFrom;
            }

            if (!string.IsNullOrWhiteSpace(receivedThrough))
            {
                payload["received_through"] = receivedThrough;
            }

            if (hasAttachments.HasValue)
            {
                payload["has_attachments"] = hasAttachments.Value;
            }

            string json = Send(
                "POST",
                "/search",
                _serializer.Serialize(payload),
                true);
            return ParseSearchResponse(json);
        }

        private SemanticSearchResponse ParseSearchResponse(string json)
        {
            Dictionary<string, object> root = Deserialize(json);
            SemanticSearchResponse response = new SemanticSearchResponse();
            response.Query = GetString(root, "query");
            response.SemanticAvailable = GetBool(root, "semantic_available");
            response.KeywordAvailable = GetBool(root, "keyword_available");
            response.CandidateCount = GetInt(root, "candidate_count");
            response.FilterApplied = GetBool(root, "filter_applied");

            object[] results = GetArray(root, "results");
            if (results != null)
            {
                for (int index = 0; index < results.Length; index++)
                {
                    Dictionary<string, object> item =
                        results[index] as Dictionary<string, object>;
                    if (item == null)
                    {
                        continue;
                    }

                    SemanticSearchResult result = new SemanticSearchResult();
                    result.MailId = GetInt(item, "mail_id");
                    result.StoreId = GetString(item, "store_id");
                    result.EntryId = GetString(item, "entry_id");
                    result.InternetMessageId =
                        GetString(item, "internet_message_id");
                    result.Subject = GetString(item, "subject");
                    result.SenderName = GetString(item, "sender_name");
                    result.SenderEmail = GetString(item, "sender_email");
                    result.ReceivedAt = GetString(item, "received_at");
                    result.FolderPath = GetString(item, "folder_path");
                    result.AttachmentNames = GetString(item, "attachment_names");
                    result.AttachmentCount = GetInt(item, "attachment_count");
                    result.Snippet = GetString(item, "snippet");
                    result.Score = GetDouble(item, "score");
                    result.SemanticRank = GetNullableInt(item, "semantic_rank");
                    result.KeywordRank = GetNullableInt(item, "keyword_rank");
                    response.Results.Add(result);
                }
            }

            return response;
        }

        private string Send(
            string method,
            string path,
            string body,
            bool requireToken)
        {
            string url = _baseUrl + path;
            HttpWebRequest request = null;
            try
            {
                request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = method;
                request.Timeout = RequestTimeoutMilliseconds;
                request.ReadWriteTimeout = RequestTimeoutMilliseconds;
                request.Proxy = null;
                request.KeepAlive = false;
                if (requireToken && !string.IsNullOrWhiteSpace(_token))
                {
                    request.Headers["Authorization"] = "Bearer " + _token;
                }

                if (body != null)
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(body);
                    request.ContentType = "application/json";
                    request.ContentLength = bytes.Length;
                    using (Stream stream = request.GetRequestStream())
                    {
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }

                using (HttpWebResponse response =
                    (HttpWebResponse)request.GetResponse())
                using (Stream stream = response.GetResponseStream())
                using (StreamReader reader = new StreamReader(
                    stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                throw new SemanticServiceException(
                    BuildErrorMessage(ex),
                    GetStatusCode(ex));
            }
            finally
            {
                if (request != null)
                {
                    request.Abort();
                }
            }
        }

        private static int GetStatusCode(WebException ex)
        {
            HttpWebResponse response = ex.Response as HttpWebResponse;
            return response == null ? 0 : (int)response.StatusCode;
        }

        private static string BuildErrorMessage(WebException ex)
        {
            if (ex.Status == WebExceptionStatus.ConnectFailure
                || ex.Status == WebExceptionStatus.Timeout
                || ex.Status == WebExceptionStatus.NameResolutionFailure)
            {
                return "本地语义搜索服务未运行。请先启动 embedding_service，"
                    + "或改用 Outlook 原生搜索。";
            }

            HttpWebResponse response = ex.Response as HttpWebResponse;
            if (response != null)
            {
                if (response.StatusCode == HttpStatusCode.Unauthorized)
                {
                    return "本地语义搜索服务拒绝了访问令牌。请重启服务以刷新令牌。";
                }

                if ((int)response.StatusCode == 503)
                {
                    return "本地语义模型尚未就绪。请先点击“下载模型”"
                        + "（或运行 embedding_service\\download-model.cmd），"
                        + "完成后再建立索引。";
                }

                try
                {
                    using (Stream stream = response.GetResponseStream())
                    using (StreamReader reader = new StreamReader(
                        stream, Encoding.UTF8))
                    {
                        string detail = reader.ReadToEnd();
                        if (!string.IsNullOrWhiteSpace(detail))
                        {
                            return "本地语义搜索服务返回错误："
                                + Truncate(detail, 200);
                        }
                    }
                }
                catch
                {
                    // Fall through to the generic message.
                }

                return "本地语义搜索服务返回错误（HTTP "
                    + (int)response.StatusCode + "）。";
            }

            return "无法连接本地语义搜索服务：" + ex.Message;
        }

        private static string Truncate(string value, int length)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= length)
            {
                return value ?? string.Empty;
            }

            return value.Substring(0, length) + "…";
        }

        private Dictionary<string, object> Deserialize(string json)
        {
            Dictionary<string, object> root =
                _serializer.DeserializeObject(json) as
                Dictionary<string, object>;
            if (root == null)
            {
                throw new SemanticServiceException(
                    "本地语义搜索服务返回了无法解析的响应。");
            }

            return root;
        }

        private static string NormalizeBaseUrl(string baseUrl)
        {
            string value = (baseUrl ?? string.Empty).Trim().TrimEnd('/');
            if (value.Length == 0)
            {
                return DefaultBaseUrl;
            }

            Uri uri;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri))
            {
                return DefaultBaseUrl;
            }

            // Only loopback is ever allowed for privacy.
            if (!uri.IsLoopback)
            {
                return DefaultBaseUrl;
            }

            return value;
        }

        private static Dictionary<string, object> GetObject(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key))
            {
                return null;
            }

            return source[key] as Dictionary<string, object>;
        }

        private static object[] GetArray(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key))
            {
                return null;
            }

            return source[key] as object[];
        }

        private static string GetString(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return string.Empty;
            }

            return Convert.ToString(source[key]);
        }

        private static bool GetBool(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(source[key]);
            }
            catch
            {
                return false;
            }
        }

        private static int GetInt(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(source[key]);
            }
            catch
            {
                return 0;
            }
        }

        private static int? GetNullableInt(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return null;
            }

            try
            {
                return Convert.ToInt32(source[key]);
            }
            catch
            {
                return null;
            }
        }

        private static long GetLong(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt64(source[key]);
            }
            catch
            {
                return 0;
            }
        }

        private static double GetDouble(
            Dictionary<string, object> source,
            string key)
        {
            if (source == null || !source.ContainsKey(key)
                || source[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToDouble(source[key]);
            }
            catch
            {
                return 0;
            }
        }
    }
}
