using System.Xml.Linq;

namespace MiniExcelLib.OpenXml.Templates;

internal class XRowInfo
{
    public string FormatText { get; set; }
    public string IEnumerablePropName { get; set; }
    public XElement Row { get; set; }
    public Type IEnumerableGenericType { get; set; }
    public IDictionary<string, MemberInfo> PropsMap { get; set; }
    public bool IsDictionary { get; set; }
    public bool IsDataTable { get; set; }
    public int CellIEnumerableValuesCount { get; set; }
    public IList<object>? CellIlListValues { get; set; }
    public IEnumerable? CellIEnumerableValues { get; set; }
    public XMergeCell? IEnumerableMercell { get; set; }
    public List<XMergeCell>? RowMercells { get; set; }
    public List<XElement>? ConditionalFormats { get; set; }
}

internal class XMergeCell
{
    public XMergeCell(XMergeCell mergeCell)
    {
        Width = mergeCell.Width;
        Height = mergeCell.Height;
        X1 = mergeCell.X1;
        Y1 = mergeCell.Y1;
        X2 = mergeCell.X2;
        Y2 = mergeCell.Y2;
        MergeCell = mergeCell.MergeCell;
    }

    public XMergeCell(XElement mergeCell)
    {
        var refAttr = mergeCell.Attribute("ref")?.Value;
        var refs = refAttr?.Split(':');

        //TODO: width,height
        var xy1 = refs[0];
        X1 = CellReferenceConverter.GetNumericalIndex(StringHelper.GetLetters(refs[0]));
        Y1 = StringHelper.GetNumber(xy1);

        var xy2 = refs[1];
        X2 = CellReferenceConverter.GetNumericalIndex(StringHelper.GetLetters(refs[1]));
        Y2 = StringHelper.GetNumber(xy2);

        Width = Math.Abs(X1 - X2) + 1;
        Height = Math.Abs(Y1 - Y2) + 1;
    }

    public XMergeCell(string x1, int y1, string x2, int y2)
    {
        X1 = CellReferenceConverter.GetNumericalIndex(x1);
        Y1 = y1;

        X2 = CellReferenceConverter.GetNumericalIndex(x2);
        Y2 = y2;

        Width = Math.Abs(X1 - X2) + 1;
        Height = Math.Abs(Y1 - Y2) + 1;
    }

    public string XY1 => $"{CellReferenceConverter.GetAlphabeticalIndex(X1)}{Y1}";
    public int X1 { get; set; }
    public int Y1 { get; set; }
    public string XY2 => $"{CellReferenceConverter.GetAlphabeticalIndex(X2)}{Y2}";
    public int X2 { get; set; }
    public int Y2 { get; set; }
    public string Ref => $"{CellReferenceConverter.GetAlphabeticalIndex(X1)}{Y1}:{CellReferenceConverter.GetAlphabeticalIndex(X2)}{Y2}";
    public XElement MergeCell { get; set; }
    public int Width { get; internal set; }
    public int Height { get; internal set; }

    public string ToXmlString(string? prefix)
        => $"<{prefix}mergeCell ref=\"{CellReferenceConverter.GetAlphabeticalIndex(X1)}{Y1}:{CellReferenceConverter.GetAlphabeticalIndex(X2)}{Y2}\"/>";
}

internal class XChildNode
{
    public string? InnerText { get; set; }
    public string ColIndex { get; set; }
    public int RowIndex { get; set; }
}

internal class MergeCellIndex(int rowStart, int rowEnd)
{
    public int RowStart { get; } = rowStart;
    public int RowEnd { get; } = rowEnd;
}


internal class MemberInfo
{
    public PropertyInfo PropertyInfo { get; set; }
    public FieldInfo FieldInfo { get; set; }
    public Type UnderlyingMemberType { get; set; }
    public PropertyInfoOrFieldInfo PropertyInfoOrFieldInfo { get; set; } = PropertyInfoOrFieldInfo.None;
}

internal class GenerateCellValuesContext
{
    public int RowIndexDiff { get; set; }
    public int HeaderDiff { get; set; }
    public string PrevHeader { get; set; }
    public string CurrentHeader { get; set; }
    public int NewRowIndex { get; set; }
    public bool IsFirst { get; set; }
    public int EnumerableIndex { get; set; }
}

internal class ConditionalFormatRange
{
    public XElement? Node { get; set; }
    public List<Range> Ranges { get; set; } = [];
}

internal struct Range
{
    public int StartColumn { get; set; }
    public int StartRow { get; set; }
    public int EndColumn { get; set; }
    public int EndRow { get; set; }

    public bool ContainsRow(int row) => StartRow <= row && row <= EndRow;
}

internal enum PropertyInfoOrFieldInfo { None, PropertyInfo, FieldInfo }

internal enum SpecialCellType { None, Group, Endgroup, Merge, Header }
