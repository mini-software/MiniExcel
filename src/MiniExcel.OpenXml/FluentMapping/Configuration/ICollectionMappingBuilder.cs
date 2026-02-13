using System.Collections;

namespace MiniExcelLib.OpenXml.FluentMapping.Configuration;

public interface ICollectionMappingBuilder<T, TCollection> where TCollection : IEnumerable
{
    ICollectionMappingBuilder<T, TCollection> StartAt(string cellAddress);
    
    ICollectionMappingBuilder<T, TCollection> WithSpacing(int spacing);
    
    ICollectionMappingBuilder<T, TCollection> WithItemMapping<TItem>(Action<IMappingConfiguration<TItem>> configure);
}