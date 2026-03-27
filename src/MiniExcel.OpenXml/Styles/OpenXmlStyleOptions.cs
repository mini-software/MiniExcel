namespace MiniExcelLib.OpenXml.Styles;

using MiniExcelLib.Core.Enums;

public class OpenXmlStyleOptions
{
    public bool WrapCellContents { get; set; }
    public OpenXmlHeaderStyle? HeaderStyle { get; set; }
    public HorizontalCellAlignment HorizontalAlignment { get; set; } = HorizontalCellAlignment.Left;
    public VerticalCellAlignment VerticalAlignment { get; set; } = VerticalCellAlignment.Bottom;
}