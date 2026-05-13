using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml.Styles;

internal class SheetStyleFormatsCache
{
    private readonly Dictionary<string, int> _formatMappings = [];
    private int _stylesCount;
    
    internal int FormatMappingsCount => _formatMappings.Count;
    internal IEnumerable<KeyValuePair<string, int>> FormatMappings => _formatMappings.Select(x => x);
    
    public void AddMappings(IEnumerable<ExcelColumnInfo> mappings)
    {
        foreach (var mapping in mappings.Where(map => map is { ExcelIgnore: false }))
        {
            if (!string.IsNullOrWhiteSpace(mapping!.ExcelFormat) && new ExcelNumberFormat(mapping.ExcelFormat).IsValid)
            {
                if (!_formatMappings.TryGetValue(mapping.ExcelFormat, out var formatId))
                {
                    formatId = _stylesCount++;
                    _formatMappings.Add(mapping.ExcelFormat, formatId);
                }

                mapping.ExcelFormatId = formatId;
            }
        }
    }
    
    internal void SetCurrentIndex(int index) => _stylesCount = index; 
}
