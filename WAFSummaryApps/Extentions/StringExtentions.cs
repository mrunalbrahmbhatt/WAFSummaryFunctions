namespace WAFSummaryApps.Extentions
{
    public static class StringExtentions
    {
        public static string ToEmptyIfNull(this string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value;
        }
        public static string ToZeroIfNullorEmpty(this string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "0" : value;
        }
    }
}
