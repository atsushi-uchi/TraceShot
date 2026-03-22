namespace TraceShot.Extensions
{
    public static class StringExtensions
    {
        public static bool In(this string value, params string[] values)
        {
            if (value == null || values == null) return false;

            return values.Contains(value, StringComparer.OrdinalIgnoreCase);
        }

        public static bool NotIn(this string value, params string[] values)
        {
            return !In(value, values);
        }
    }
}
