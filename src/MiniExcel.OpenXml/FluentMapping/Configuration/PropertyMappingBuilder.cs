using System.Text.RegularExpressions;

namespace MiniExcelLib.OpenXml.FluentMapping.Configuration;

internal partial class PropertyMappingBuilder<T, TProperty> : IPropertyMappingBuilder<T, TProperty>
{
#if NET7_0_OR_GREATER
    [GeneratedRegex("^[A-Z]+[0-9]+$")] private static partial Regex CellAddressRegexImpl();
    private static readonly Regex CellAddressRegex = CellAddressRegexImpl();
#else
    private static readonly Regex CellAddressRegex = new("^[A-Z]+[0-9]+$", RegexOptions.Compiled);
#endif
    
    private readonly PropertyMapping _mapping;
    
    internal PropertyMappingBuilder(PropertyMapping mapping)
    {
        _mapping = mapping;
    }
    
    public IPropertyMappingBuilder<T, TProperty> ToCell(string cellAddress)
    {
        if (string.IsNullOrEmpty(cellAddress))
            throw new ArgumentException("Cell address cannot be null or empty", nameof(cellAddress));
        
        // Basic validation for cell address format (e.g., A1, AB123, etc.)
        if (!CellAddressRegex.IsMatch(cellAddress))
            throw new ArgumentException($"Invalid cell address format: {cellAddress}. Expected format like A1, B2, AA10, etc.", nameof(cellAddress));
            
        _mapping.CellAddress = cellAddress;
        return this;
    }
    
    public IPropertyMappingBuilder<T, TProperty> WithFormat(string format)
    {
        _mapping.Format = format;
        return this;
    }
    
    public IPropertyMappingBuilder<T, TProperty> WithFormula(string formula)
    {
        _mapping.Formula = formula;
        return this;
    }
}