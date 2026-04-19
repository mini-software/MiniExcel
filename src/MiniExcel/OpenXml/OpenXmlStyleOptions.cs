using System.Drawing;

namespace MiniExcelLibs.OpenXml;

public class OpenXmlStyleOptions
{
    public OpenXmlHeaderStyle HeaderStyle { get; set; }
    public bool WrapCellContents { get; set; }
    public HorizontalCellAlignment HorizontalAlignment { get; set; }
    public VerticalCellAlignment VerticalAlignment { get; set; }
}

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
    
public enum HorizontalCellAlignment { Left, Center, Right }
public enum VerticalCellAlignment { Bottom, Center, Top }