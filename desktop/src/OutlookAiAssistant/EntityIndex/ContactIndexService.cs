using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using OutlookAiAssistant.Configuration;

namespace OutlookAiAssistant.EntityIndex
{
    /// <summary>
    /// Rebuilds the index only when explicitly called by the Settings button.
    /// No timer, startup hook or Outlook event invokes this service.
    /// </summary>
    public sealed class ContactIndexService
    {
        private readonly OutlookContactScanner _outlookScanner;
        private readonly ContactProfileExtractor _extractor;
        private readonly ContactIndexStore _store;

        internal ContactIndexService(
            OutlookContactScanner outlookScanner,
            ContactProfileExtractor extractor,
            ContactIndexStore store)
        {
            _outlookScanner = outlookScanner;
            _extractor = extractor;
            _store = store;
        }

        public async Task<ContactIndex> RebuildAsync(
            AppSettings settings,
            CancellationToken cancellationToken,
            Action<int, int> progress)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            cancellationToken.ThrowIfCancellationRequested();
            ContactDiscoveryResult discovery =
                _outlookScanner.Discover(settings);
            IList<ContactCandidate> candidates = discovery.Candidates;
            ContactIndex previous = TryLoadPrevious();
            if (candidates.Count == 0 && previous.Contacts.Count > 0)
            {
                throw new InvalidOperationException(
                    "本次未发现任何联系人，旧联系人索引已保留。");
            }

            Dictionary<string, ContactProfile> previousByEmail =
                IndexByEmail(previous.Contacts);
            List<ContactProfile> contacts = new List<ContactProfile>();
            int extractionAttempts = 0;
            int extractionSuccesses = 0;

            for (int index = 0; index < candidates.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (progress != null)
                {
                    progress(index, candidates.Count);
                }

                ContactCandidate candidate = candidates[index];
                IList<ContactRepresentativeMessage> messages =
                    _outlookScanner.LoadRepresentativeMessages(candidate);
                ContactProfile profile = null;
                if (messages.Count > 0)
                {
                    extractionAttempts++;
                    try
                    {
                        profile = await _extractor.ExtractAsync(
                            candidate,
                            messages,
                            settings,
                            cancellationToken);
                        extractionSuccesses++;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch
                    {
                        previousByEmail.TryGetValue(candidate.Email, out profile);
                    }
                }

                contacts.Add(profile
                    ?? ContactProfileExtractor.CreateBaseline(candidate));
            }

            if (extractionAttempts > 0 && extractionSuccesses == 0)
            {
                throw new InvalidOperationException(
                    "联系人资料提取全部失败，旧联系人索引已保留。请检查 API 设置和网络。");
            }

            contacts.Sort(delegate(ContactProfile left, ContactProfile right)
            {
                return string.Compare(
                    left.DisplayName,
                    right.DisplayName,
                    StringComparison.OrdinalIgnoreCase);
            });

            ContactIndex rebuilt = new ContactIndex();
            rebuilt.CreatedAt = DateTime.UtcNow.ToString("o");
            rebuilt.Contacts.AddRange(contacts);
            rebuilt.FolderContext.AddRange(discovery.FolderContext);
            _store.Save(rebuilt);
            if (progress != null)
            {
                progress(candidates.Count, candidates.Count);
            }

            return rebuilt;
        }

        private ContactIndex TryLoadPrevious()
        {
            try
            {
                return _store.Load();
            }
            catch
            {
                return new ContactIndex();
            }
        }

        private static Dictionary<string, ContactProfile> IndexByEmail(
            IList<ContactProfile> contacts)
        {
            Dictionary<string, ContactProfile> result =
                new Dictionary<string, ContactProfile>(
                    StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < contacts.Count; index++)
            {
                ContactProfile contact = contacts[index];
                if (contact != null && !string.IsNullOrWhiteSpace(contact.Email))
                {
                    result[contact.Email] = contact;
                }
            }

            return result;
        }
    }
}
