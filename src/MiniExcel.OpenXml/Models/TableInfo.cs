namespace MiniExcelLib.OpenXml.Models;

public class TableInfo
{
    internal TableInfo(string name, IEnumerable<string> columns, string? referenceCells, bool hiddenHeader)
    {
        Name = name;
        Columns = [..columns];
        ReferenceCells = referenceCells;
        HiddenHeader = hiddenHeader;
    }

    public string Name { get; private set; }
    public string[] Columns { get; private set; }
    public string? ReferenceCells { get; private set; }
    public bool HiddenHeader { get; private set; }
}
