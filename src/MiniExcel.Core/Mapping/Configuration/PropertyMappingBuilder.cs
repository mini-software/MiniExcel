namespace MiniExcelLib.Core.Mapping.Configuration;

internal class PropertyMappingBuilder<T, TProperty> : IPropertyMappingBuilder<T, TProperty>
{
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
        if (!Regex.IsMatch(cellAddress, @"^[A-Z]+[0-9]+$"))
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