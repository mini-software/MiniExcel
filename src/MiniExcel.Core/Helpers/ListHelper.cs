namespace MiniExcelLib.Core.Helpers;

internal static class ListHelper
{
    internal static bool StartsWith<T>(this IList<T> span, IList<T> value) where T : IEquatable<T>
    {
        if (value is [])
            return true;

        if (span.Count < value.Count)
            return false;

        return span.Take(value.Count).SequenceEqual(value);
    }
}