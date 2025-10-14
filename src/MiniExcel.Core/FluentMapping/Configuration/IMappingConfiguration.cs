using System.Linq.Expressions;

namespace MiniExcelLib.Core.FluentMapping.Configuration;

public interface IMappingConfiguration<T>
{
    IPropertyMappingBuilder<T, TProperty> Property<TProperty>(Expression<Func<T, TProperty>> property);
    ICollectionMappingBuilder<T, TCollection> Collection<TCollection>(Expression<Func<T, TCollection>> collection) where TCollection : IEnumerable;
    IMappingConfiguration<T> ToWorksheet(string worksheetName);
}