namespace MiniExcelLib.OpenXml.FluentMapping.Configuration;

public interface IPropertyMappingBuilder<T, TProperty>
{
    IPropertyMappingBuilder<T, TProperty> ToCell(string cellAddress);
    IPropertyMappingBuilder<T, TProperty> WithFormat(string format);
    IPropertyMappingBuilder<T, TProperty> WithFormula(string formula);
}