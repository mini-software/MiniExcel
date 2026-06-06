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
            cell.TryAdd(key, null);
        }

        return cell;
    }

    public static IDictionary<string, object?> CreateEmptyByHeaders(Dictionary<int, string> headers)
    {
        IDictionary<string, object?> cell = new ExpandoObject();
        foreach (var hr in headers)
        {
            cell.TryAdd(hr.Value, null);
        }

        return cell;
    }
}
