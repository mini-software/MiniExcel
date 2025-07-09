namespace MiniExcelLib.OpenXml.Models;

public class ExcelRangeElement
{
    internal ExcelRangeElement(int startIndex, int endIndex)
    {
        if (startIndex > endIndex)
            throw new ArgumentException("StartIndex value cannot be greater than EndIndex value.");

        StartIndex = startIndex; 
        EndIndex = endIndex;
    }

    public int StartIndex { get; }
    public int EndIndex { get; }

    public int Count => EndIndex - StartIndex + 1;
}

public class ExcelRange(int maxRow, int maxColumn)
{
    public string StartCell { get; internal set; }
    public string EndCell { get; internal set; }

    public ExcelRangeElement Rows { get; } = new(1, maxRow);
    public ExcelRangeElement Columns { get; } = new(1, maxColumn);
}