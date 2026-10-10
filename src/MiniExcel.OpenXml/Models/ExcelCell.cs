namespace MiniExcelLib.OpenXml.Models;

/// <summary>
/// A worksheet cell together with the metadata that row queries leave out.
/// </summary>
public class ExcelCell
{
    internal ExcelCell(string reference, int rowIndex, int columnIndex, object? value, bool hasFormula, string? formula, ExcelColor? fontColor)
    {
        Reference = reference;
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Value = value;
        HasFormula = hasFormula;
        Formula = formula;
        FontColor = fontColor;
    }

    /// <summary>
    /// The A1-style reference of the cell, e.g. "B3".
    /// </summary>
    public string Reference { get; }

    /// <summary>
    /// The 1-based row number of the cell.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// The 1-based column number of the cell.
    /// </summary>
    public int ColumnIndex { get; }

    /// <summary>
    /// The cell value, converted the same way as in row queries. For formula cells this is the result cached by the application that last saved the file.
    /// </summary>
    public object? Value { get; }

    /// <summary>
    /// True when the value of the cell is computed by a formula.
    /// </summary>
    public bool HasFormula { get; }

    /// <summary>
    /// The formula text as stored in the file, without the leading '='.
    /// Cells that only share the formula of another cell report <see cref="HasFormula"/> as true and this property as <see langword="null"/>.
    /// </summary>
    public string? Formula { get; }

    /// <summary>
    /// The font color applied to the cell, or <see langword="null"/> when its font declares no color.
    /// </summary>
    public ExcelColor? FontColor { get; }
}

/// <summary>
/// A color as declared in the workbook styles. Exactly one of <see cref="Auto"/>, <see cref="Rgb"/>, <see cref="Theme"/> or <see cref="Indexed"/> is normally set.
/// </summary>
public class ExcelColor
{
    internal ExcelColor(bool auto, string? rgb, int? theme, int? indexed, double tint)
    {
        Auto = auto;
        Rgb = rgb;
        Theme = theme;
        Indexed = indexed;
        Tint = tint;
    }

    /// <summary>
    /// True when the color is the automatic system color.
    /// </summary>
    public bool Auto { get; }

    /// <summary>
    /// The ARGB hex value, e.g. "FF000000".
    /// </summary>
    public string? Rgb { get; }

    /// <summary>
    /// The zero-based index into the workbook theme colors (0 = Light 1, 1 = Dark 1, 2 = Light 2, 3 = Dark 2, 4-9 = Accent 1-6).
    /// </summary>
    public int? Theme { get; }

    /// <summary>
    /// The index into the legacy indexed color palette.
    /// </summary>
    public int? Indexed { get; }

    /// <summary>
    /// The lightening (positive) or darkening (negative) applied to the color, from -1 to 1.
    /// </summary>
    public double Tint { get; }
}
