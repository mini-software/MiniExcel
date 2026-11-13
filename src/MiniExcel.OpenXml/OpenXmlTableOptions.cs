namespace MiniExcelLib.OpenXml;

/// <summary>
/// Exports each sheet's data as an Excel table. Set it on <see cref="OpenXmlConfiguration.Table"/>.
/// </summary>
/// <remarks>
/// The table covers the header row and the data, so columns must be adjacent. Combine it with
/// <see cref="TableStyles.None"/> to let the table style alone format the cells.
/// </remarks>
public class OpenXmlTableOptions
{
    /// <summary>
    /// The table name. The first sheet's table gets this name and later sheets append their sheet number.
    /// Defaults to Table1, Table2 and so on.
    /// </summary>
    public string? Name { get; set; }

    /// <summary>
    /// A built-in table style such as TableStyleMedium2 or TableStyleLight9.
    /// </summary>
    public string Style { get; set; } = "TableStyleMedium2";

    /// <summary>
    /// Shows filter buttons in the header row.
    /// </summary>
    public bool ShowFilterButton { get; set; } = true;

    /// <summary>
    /// Shades alternate rows.
    /// </summary>
    public bool ShowRowStripes { get; set; } = true;

    /// <summary>
    /// Shades alternate columns.
    /// </summary>
    public bool ShowColumnStripes { get; set; }

    /// <summary>
    /// Highlights the first column.
    /// </summary>
    public bool ShowFirstColumn { get; set; }

    /// <summary>
    /// Highlights the last column.
    /// </summary>
    public bool ShowLastColumn { get; set; }
}
