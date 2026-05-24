namespace MiniExcelLib.OpenXml.Styles.Builder;

// todo: find a way to make it compatible with SharedStringsDiskCache
internal class SheetStyleFormatsCache
{
    private readonly Dictionary<string, int> _formatMappings = [];
    private int _stylesCount;
    
    internal int FormatMappingsCount => _formatMappings.Count;
    internal IEnumerable<(string Format, int FormatId)> FormatMappings => _formatMappings.Select(x => (x.Key, x.Value));
    
    public void AddMappings(IEnumerable<MiniExcelColumnMapping?> mappings)
    {
        foreach (var mapping in mappings.Where(map => map is { ExcelIgnoreColumn: false }))
        {
            if (!string.IsNullOrWhiteSpace(mapping!.ExcelFormat) && new OpenXmlNumberFormatHelper(mapping.ExcelFormat).IsValid)
            {
                if (!_formatMappings.TryGetValue(mapping.ExcelFormat, out var formatId))
                {
                    formatId = _stylesCount++;
                    _formatMappings.Add(mapping.ExcelFormat, formatId);
                }

                mapping.SetFormatId(formatId);
            }
        }
    }
    
    internal void SetCurrentIndex(int index) => _stylesCount = index; 
}
