using DocumentFormat.OpenXml.Spreadsheet;

namespace MiniExcelLibs.Benchmarks.Utils;

internal static class Extensions
{
    internal static void Add(this Row row, params string[] values)
    {
        foreach (var value in values)
        {
            var cell = new Cell
            {
                CellValue = new CellValue(value),
                DataType = CellValues.String
            };
            
            row.Append(cell);
        }
    }
}
