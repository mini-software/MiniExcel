using System.Dynamic;

namespace MiniExcelLib.Core.Helpers;

public static class ExpandoHelper
{
    public static IDictionary<string, object?> CreateEmptyByIndices(int maxColumnIndex, int startCellIndex)
    {
        IDictionary<string, object?> cell = new ExpandoObject();
        for (int i = startCellIndex; i <= maxColumnIndex; i++)
        {
            var key = CellReferenceConverter.GetAlphabeticalIndex(i);
#if NETCOREAPP2_0_OR_GREATER
            cell.TryAdd(key, null);
#else
            if (!cell.ContainsKey(key))
                cell.Add(key, null);
#endif
        }

        return cell;
    }

    public static IDictionary<string, object?> CreateEmptyByHeaders(Dictionary<int, string> headers)
    {
        IDictionary<string, object?> cell = new ExpandoObject();
        foreach (var hr in headers)
        {
#if NETCOREAPP2_0_OR_GREATER
            cell.TryAdd(hr.Value, null);
#else
            if (!cell.ContainsKey(hr.Value))
                cell.Add(hr.Value, null);
#endif
        }

        return cell;
    }
}
