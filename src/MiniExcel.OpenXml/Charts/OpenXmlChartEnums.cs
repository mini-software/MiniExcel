namespace MiniExcelLib.OpenXml.Charts;

public enum OpenXmlChartType
{
    /// <summary>Points at X and Y values, drawn as markers.</summary>
    Scatter,
    /// <summary>Points at X and Y values, joined by straight lines without markers.</summary>
    ScatterLines,
    /// <summary>Values at evenly spaced categories, joined by straight lines without markers.</summary>
    Line,
    /// <summary>Vertical bars side by side.</summary>
    Column,
    /// <summary>Vertical bars stacked on top of each other.</summary>
    StackedColumn,
    /// <summary>Horizontal bars side by side.</summary>
    Bar,
    /// <summary>Horizontal bars stacked end to end.</summary>
    StackedBar
}

public enum OpenXmlChartLegendPosition
{
    Right,
    Left,
    Top,
    Bottom,
    TopRight
}

public enum OpenXmlChartTickMark
{
    None,
    Inside,
    Outside,
    Cross
}
