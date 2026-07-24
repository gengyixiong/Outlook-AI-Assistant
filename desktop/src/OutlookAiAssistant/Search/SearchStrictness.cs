namespace OutlookAiAssistant.Search
{
    /// <summary>
    /// Controls how many allow-listed SearchPlan conditions are included.
    /// Broad search favors recall; precise search adds contextual hints.
    /// </summary>
    public enum SearchStrictness
    {
        Broad,
        Recommended,
        Precise
    }
}
