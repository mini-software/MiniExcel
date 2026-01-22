using System.Drawing;
using MiniExcelLib.Core.Enums;

namespace MiniExcelLib.OpenXml.Styles;

public class OpenXmlHeaderStyle
{
    /// <summary>
    /// Whether to wrap the content of the header
    /// </summary>
    public bool WrapText { get; set; }

    /// <summary>
    /// The RGB background color in the filtered state
    /// </summary>
    public Color BackgroundColor { get; set; } = Color.FromArgb(0x284472C4);

    /// <summary>
    /// Horizontal alignment
    /// </summary>
    public HorizontalCellAlignment HorizontalAlignment { get; set; } = HorizontalCellAlignment.Left;

    /// <summary>
    /// Vertical alignment
    /// </summary>
    public VerticalCellAlignment VerticalAlignment { get; set; } = VerticalCellAlignment.Bottom;
}