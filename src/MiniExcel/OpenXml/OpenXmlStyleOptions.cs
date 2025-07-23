using MiniExcelLibs.OpenXml.Enums;

namespace MiniExcelLibs.OpenXml;

public class OpenXmlStyleOptions
{
    public bool WrapCellContents { get; set; }

    public OpenXmlColumnStyle ColumnStyle { get; set; }

}

public class OpenXmlColumnStyle
{
    /// <summary>
    /// 过滤状态下的RGB背景色
    /// </summary>
    public string FilterBackgroundColor { get; set; } 

    /// <summary>
    /// 水平对齐
    /// </summary>
    public CellHorizontalType Horizontal { get; set; } = CellHorizontalType.left;

    /// <summary>
    /// 垂直对齐
    /// </summary>
    public CellVerticalType Vertical { get; set; } = CellVerticalType.bottom;

    /// <summary>
    /// 是否自动换行
    /// </summary>
    public bool WrapText { get; set; }

}