using System.Drawing;
// A cell format added on top of the builder's own. CustomNumberFormat and FontColor are positions in
// NumberFormats and FontColors, or null; BuiltInNumberFormatId applies when there is no custom number format.
using CustomCellFormat = (int? CustomNumberFormat, int BuiltInNumberFormatId, int? FontColor);

namespace MiniExcelLib.OpenXml.Styles.Builder;

// todo: find a way to make it compatible with SharedStringsDiskCache
internal class SheetStyleFormatsCache
{
    private readonly Dictionary<string, int> _numberFormatIndexes = [];
    private readonly List<string> _numberFormats = [];
    private readonly Dictionary<int, int> _fontColorIndexes = [];
    private readonly List<Color> _fontColors = [];
    private readonly Dictionary<CustomCellFormat, int> _cellFormatIndexes = [];
    private readonly List<CustomCellFormat> _cellFormats = [];
    private int _firstCellFormatIndex;

    internal IReadOnlyList<string> NumberFormats => _numberFormats;
    internal IReadOnlyList<Color> FontColors => _fontColors;
    internal IReadOnlyList<CustomCellFormat> CellFormats => _cellFormats;

    public void AddMappings(IEnumerable<MiniExcelColumnMapping?> mappings)
    {
        foreach (var mapping in mappings.Where(map => map is { ExcelIgnoreColumn: false }))
        {
            if (!string.IsNullOrWhiteSpace(mapping!.ExcelFormat) && new OpenXmlNumberFormatHelper(mapping.ExcelFormat).IsValid)
            {
                mapping.SetFormatId(GetCellFormatIndex(mapping.ExcelFormat, 0, null));
            }
        }
    }

    /// <summary>
    /// Gets the cellXfs index of a format, adding the format when it is new.
    /// </summary>
    /// <param name="numberFormat">A valid custom number format, or null.</param>
    /// <param name="builtInNumberFormatId">The built-in numFmtId used when <paramref name="numberFormat"/> is null.</param>
    /// <param name="fontColor">The font color, or null for the default font.</param>
    internal int GetCellFormatIndex(string? numberFormat, int builtInNumberFormatId, Color? fontColor)
    {
        int? customNumberFormat = null;
        if (numberFormat is not null)
        {
            if (!_numberFormatIndexes.TryGetValue(numberFormat, out var position))
            {
                position = _numberFormats.Count;
                _numberFormats.Add(numberFormat);
                _numberFormatIndexes.Add(numberFormat, position);
            }

            customNumberFormat = position;
            builtInNumberFormatId = 0;
        }

        int? fontColorPosition = null;
        if (fontColor is { } color)
        {
            var argb = color.ToArgb();
            if (!_fontColorIndexes.TryGetValue(argb, out var position))
            {
                position = _fontColors.Count;
                _fontColors.Add(color);
                _fontColorIndexes.Add(argb, position);
            }

            fontColorPosition = position;
        }

        CustomCellFormat format = (customNumberFormat, builtInNumberFormatId, fontColorPosition);
        if (!_cellFormatIndexes.TryGetValue(format, out var index))
        {
            index = _firstCellFormatIndex + _cellFormats.Count;
            _cellFormats.Add(format);
            _cellFormatIndexes.Add(format, index);
        }

        return index;
    }

    internal void SetCurrentIndex(int index) => _firstCellFormatIndex = index;
}
