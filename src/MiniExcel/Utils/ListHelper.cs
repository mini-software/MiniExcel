using System;
using System.Collections.Generic;
using System.Linq;

namespace MiniExcelLibs.Utils;

internal static class IEnumerableHelper
{
    internal static bool StartsWith<T>(this IList<T> span, IList<T> value) where T : IEquatable<T>
    {
        if (value is [])
            return true;

        var b = span.Take(value.Count);
        var bCount = b.Count();
        if (bCount != value.Count)
            return false;

        for (int i = 0; i < bCount; i++)
            if (!span[i].Equals(value[i]))
                return false;

        return true;
    }
}