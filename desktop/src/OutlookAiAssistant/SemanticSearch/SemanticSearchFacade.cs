using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAiAssistant.Diagnostics;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>
    /// UI-facing facade over the local semantic service. It translates the
    /// service's structured response into display rows and resolves each row
    /// to a real Outlook item on demand.
    ///
    /// It never throws for service outages: callers receive a response with a
    /// user-facing message instead, so the add-in cannot crash when the local
    /// service is not running.
    /// </summary>
    public sealed class SemanticSearchFacade
    {
        private readonly SemanticSearchClient _client;
        private readonly OutlookMailLocator _locator;

        public SemanticSearchFacade(
            SemanticSearchClient client,
            OutlookMailLocator locator)
        {
            if (client == null)
            {
                throw new ArgumentNullException("client");
            }

            _client = client;
            _locator = locator;
        }

        /// <summary>Combined outcome of a semantic search attempt.</summary>
        public sealed class Outcome
        {
            public bool Success { get; set; }
            public string Message { get; set; }
            public SemanticSearchResponse Response { get; set; }

            public Outcome()
            {
                Message = string.Empty;
            }
        }

        public Task<Outcome> SearchAsync(
            string description,
            CancellationToken cancellationToken)
        {
            return Task.Factory.StartNew(
                delegate { return Search(description); },
                cancellationToken);
        }

        public Outcome Search(string description)
        {
            Outcome outcome = new Outcome();
            try
            {
                SemanticSearchResponse response = _client.Search(
                    description, null, null, null, null, null);
                outcome.Success = true;
                outcome.Response = response;
                if (response.Results.Count == 0)
                {
                    outcome.Message = response.FilterApplied
                        ? "没有找到匹配的邮件（已应用筛选条件）。"
                        : "没有找到匹配的邮件。可以尝试更宽泛的描述，"
                            + "或先建立语义索引。";
                }
                else if (!response.SemanticAvailable)
                {
                    outcome.Message = "语义模型当前不可用，已仅使用关键词匹配。"
                        + "请检查本地语义服务中的模型状态。";
                }

                return outcome;
            }
            catch (SemanticServiceException ex)
            {
                outcome.Success = false;
                outcome.Message = ex.Message;
                return outcome;
            }
            catch (Exception ex)
            {
                // Defensive: the UI must never crash on an unexpected error.
                Logger.Error("Semantic search failed.", ex);
                outcome.Success = false;
                outcome.Message = "语义搜索失败：" + ex.Message;
                return outcome;
            }
        }

        /// <summary>
        /// Opens the Outlook item behind one result. Returns null on success,
        /// otherwise a user-facing message.
        /// </summary>
        public string OpenResult(SemanticSearchResult result)
        {
            try
            {
                return _locator.OpenMail(result);
            }
            catch (Exception ex)
            {
                Logger.Error("Opening a search result failed.", ex);
                return "无法打开邮件：" + ex.Message;
            }
        }
    }
}
