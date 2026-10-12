using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAiAssistant.Configuration;
using OutlookAiAssistant.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAiAssistant.SemanticSearch
{
    /// <summary>Progress reported during a semantic index build.</summary>
    public sealed class SemanticIndexProgress
    {
        public int Processed { get; set; }
        public int Total { get; set; }
        public int Indexed { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public string CurrentSubject { get; set; }

        public SemanticIndexProgress()
        {
            CurrentSubject = string.Empty;
        }
    }

    /// <summary>Outcome of one semantic index build or sync.</summary>
    public sealed class SemanticIndexReport
    {
        public int Candidates { get; set; }
        public int Indexed { get; set; }
        public int Skipped { get; set; }
        public int Failed { get; set; }
        public bool Cancelled { get; set; }
        public string Error { get; set; }

        public SemanticIndexReport()
        {
            Error = string.Empty;
        }
    }

    /// <summary>
    /// Orchestrates a semantic index build. It enumerates candidate mail
    /// metadata on the UI thread, then sends each mail to the local service
    /// one at a time, reporting progress and honouring cancellation.
    ///
    /// The build is deliberately driven by the UI thread for COM access and
    /// yields to the message loop between mails so Outlook stays responsive.
    /// </summary>
    public sealed class SemanticIndexService
    {
        private readonly OutlookMailScanner _scanner;
        private readonly SemanticSearchClient _client;

        public SemanticIndexService(
            Outlook.Application application,
            SemanticSearchClient client)
        {
            if (application == null)
            {
                throw new ArgumentNullException("application");
            }

            if (client == null)
            {
                throw new ArgumentNullException("client");
            }

            _scanner = new OutlookMailScanner(application);
            _client = client;
        }

        /// <summary>
        /// Runs the index build on a background thread while marshalling every
        /// COM call back to the UI thread the form supplies.
        /// </summary>
        public async Task<SemanticIndexReport> BuildIndexAsync(
            AppSettings settings,
            CancellationToken cancellationToken,
            Action<SemanticIndexProgress> progress,
            Func<Func<object>, object> uiInvoke)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            SemanticIndexReport report = new SemanticIndexReport();
            return await Task.Factory.StartNew(
                delegate
                {
                    try
                    {
                        RunBuild(
                            settings,
                            cancellationToken,
                            progress,
                            uiInvoke,
                            report);
                    }
                    catch (OperationCanceledException)
                    {
                        report.Cancelled = true;
                    }
                    catch (Exception ex)
                    {
                        report.Error = ex.Message;
                        Logger.Error("Semantic index build failed.", ex);
                    }

                    return report;
                },
                cancellationToken,
                TaskCreationOptions.LongRunning,
                TaskScheduler.Default);
        }

        private void RunBuild(
            AppSettings settings,
            CancellationToken cancellationToken,
            Action<SemanticIndexProgress> progress,
            Func<Func<object>, object> uiInvoke,
            SemanticIndexReport report)
        {
            cancellationToken.ThrowIfCancellationRequested();

            IList<MailCandidate> candidates = (IList<MailCandidate>)uiInvoke(
                delegate
                {
                    return _scanner.EnumerateCandidates(
                        settings,
                        delegate(int count)
                        {
                            if (progress != null)
                            {
                                progress(new SemanticIndexProgress
                                {
                                    Total = count,
                                    CurrentSubject = "正在枚举邮件…"
                                });
                            }
                        });
                });

            report.Candidates = candidates.Count;
            string ownerMailbox = (string)uiInvoke(
                delegate { return _scanner.GetOwnerMailbox(); });

            SemanticIndexProgress state = new SemanticIndexProgress();
            state.Total = candidates.Count;
            if (progress != null)
            {
                progress(Clone(state));
            }

            for (int index = 0; index < candidates.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                MailCandidate candidate = candidates[index];
                state.CurrentSubject = candidate.Subject;

                // Only read the MailItem body when the mail may have changed.
                // The service decides whether the mail is unchanged, but we
                // must still send the body to let it compute the hash.
                MailIndexPayload payload = (MailIndexPayload)uiInvoke(
                    delegate
                    {
                        return _scanner.ReadMailPayload(
                            candidate, ownerMailbox);
                    });

                if (payload == null)
                {
                    state.Failed++;
                    state.Processed++;
                    report.Failed++;
                    if (progress != null)
                    {
                        progress(Clone(state));
                    }

                    continue;
                }

                try
                {
                    string response = _client
                        .IndexMailAsync(payload, cancellationToken)
                        .GetAwaiter()
                        .GetResult();
                    bool skipped = response.IndexOf(
                        "\"skipped\":true",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                    if (skipped)
                    {
                        state.Skipped++;
                        report.Skipped++;
                    }
                    else
                    {
                        state.Indexed++;
                        report.Indexed++;
                    }
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    // A single mail failure must not abort the whole build.
                    state.Failed++;
                    report.Failed++;
                    Logger.Error(
                        "Indexing one mail failed; continuing.", ex);
                }

                state.Processed++;
                if (progress != null)
                {
                    progress(Clone(state));
                }
            }

            _client.ReportSync(null);
        }

        public void DeleteMail(string storeId, string entryId)
        {
            _client.DeleteMail(storeId, entryId);
        }

        public void ClearIndex()
        {
            _client.ClearIndex();
        }

        private static SemanticIndexProgress Clone(SemanticIndexProgress source)
        {
            return new SemanticIndexProgress
            {
                Processed = source.Processed,
                Total = source.Total,
                Indexed = source.Indexed,
                Skipped = source.Skipped,
                Failed = source.Failed,
                CurrentSubject = source.CurrentSubject
            };
        }
    }
}
