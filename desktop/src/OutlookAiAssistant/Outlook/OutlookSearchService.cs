using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.OutlookIntegration
{
    /// <summary>
    /// Executes AQS through Outlook Instant Search. No message content is read or
    /// uploaded by this service; Outlook owns indexing and result rendering.
    /// </summary>
    public sealed class OutlookSearchService
    {
        private readonly Outlook.Application _application;

        public OutlookSearchService(Outlook.Application application)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            _application = application;
        }

        public void Search(string aqsQuery, string scope)
        {
            if (string.IsNullOrWhiteSpace(aqsQuery))
            {
                throw new ArgumentException("搜索条件不能为空。", "aqsQuery");
            }

            Outlook.Explorer explorer = null;
            try
            {
                explorer = _application.ActiveExplorer();
                if (explorer == null)
                {
                    throw new InvalidOperationException(
                        "请先打开 Outlook 主邮件列表窗口再执行搜索。");
                }

                Outlook.OlSearchScope outlookScope =
                    string.Equals(scope, "current_folder", StringComparison.OrdinalIgnoreCase)
                        ? Outlook.OlSearchScope.olSearchScopeCurrentFolder
                        : Outlook.OlSearchScope.olSearchScopeAllFolders;
                explorer.Search(aqsQuery, outlookScope);
            }
            finally
            {
                ComRelease.Release(explorer);
            }
        }
    }
}

