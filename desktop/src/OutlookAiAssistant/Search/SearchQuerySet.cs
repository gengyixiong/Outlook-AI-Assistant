using System;

namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// AQS variants compiled from one validated SearchPlan. These strings are
    /// produced locally and can be reused without another AI request.
    /// </summary>
    public sealed class SearchQuerySet
    {
        public string Broad { get; private set; }
        public string Recommended { get; private set; }
        public string Precise { get; private set; }

        public SearchQuerySet(
            string broad,
            string recommended,
            string precise)
        {
            Broad = broad ?? string.Empty;
            Recommended = recommended ?? string.Empty;
            Precise = precise ?? string.Empty;
        }

        public string Get(SearchStrictness strictness)
        {
            switch (strictness)
            {
                case SearchStrictness.Broad:
                    return Broad;
                case SearchStrictness.Recommended:
                    return Recommended;
                case SearchStrictness.Precise:
                    return Precise;
                default:
                    throw new ArgumentOutOfRangeException("strictness");
            }
        }
    }
}
