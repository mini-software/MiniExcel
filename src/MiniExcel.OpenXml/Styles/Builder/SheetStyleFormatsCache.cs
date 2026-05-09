namespace MiniExcelLib.OpenXml.Styles.Builder;

internal class SheetStyleFormatsCache
{
    private readonly Dictionary<string, int> _formatMappings = [];
    private int _currentIndex;
    
    internal int FormatMappingsCount => _formatMappings.Count;
    internal IEnumerable<(string Format, int FormatId)> FormatMappings => _formatMappings.Select(x => (x.Key, x.Value));
    
    public void AddMappings(IEnumerable<MiniExcelColumnMapping?> mappings)
    {
        foreach (var mapping in mappings)
        {
            if (!string.IsNullOrWhiteSpace(mapping?.ExcelFormat) && new OpenXmlNumberFormatHelper(mapping.ExcelFormat).IsValid)
            {
                if (!_formatMappings.TryGetValue(mapping.ExcelFormat, out var formatId))
                {
                    formatId = _currentIndex;
                    _formatMappings.Add(mapping.ExcelFormat, formatId);
                }

                mapping.SetFormatId(formatId);
                _currentIndex++;
            }
        }
    }
    
    internal void SetCurrentIndex(int index) => _currentIndex = index; 
}
