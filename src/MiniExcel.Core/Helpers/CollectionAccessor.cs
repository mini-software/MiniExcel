using System.Linq.Expressions;

namespace MiniExcelLib.Core.Helpers;

/// <summary>
/// Optimized collection access utilities to reduce code duplication across mapping components.
/// Provides consistent handling of IList vs IEnumerable patterns.
/// </summary>
internal static class CollectionAccessor
{
    /// <summary>
    /// Gets an item at the specified offset from a collection, with optimized handling for IList.
    /// </summary>
    /// <param name="enumerable">The collection to access</param>
    /// <param name="offset">The zero-based index of the item to retrieve</param>
    /// <returns>The item at the specified offset, or null if not found or out of bounds</returns>
    public static object? GetItemAt(IEnumerable? enumerable, int offset)
    {
        return enumerable switch
        {
            null => null,
            IList list => offset < list.Count ? list[offset] : null,
            _ => enumerable.Cast<object>().Skip(offset).FirstOrDefault()
        };
    }

    /// <summary>
    /// Creates a typed list of the specified item type.
    /// </summary>
    /// <param name="itemType">The type of items the list will contain</param>
    /// <returns>A new generic List instance</returns>
    public static IList CreateTypedList(Type itemType)
    {
        var listType = typeof(List<>).MakeGenericType(itemType);
        return (IList)Activator.CreateInstance(listType)!;
    }

    /// <summary>
    /// Converts a generic collection to the appropriate collection type (array or list).
    /// </summary>
    /// <param name="list">The source list to convert</param>
    /// <param name="targetType">The target collection type</param>
    /// <param name="itemType">The type of items in the collection</param>
    /// <returns>The converted collection</returns>
    public static object FinalizeCollection(IList list, Type targetType, Type itemType)
    {
        if (!targetType.IsArray) 
            return list;
        
        var array = Array.CreateInstance(itemType, list.Count);
        list.CopyTo(array, 0);
        return array;

    }

    /// <summary>
    /// Creates a default item factory for the specified type.
    /// </summary>
    /// <param name="itemType">The type to create instances of</param>
    /// <returns>A factory function that creates new instances</returns>
    public static Func<object?> CreateItemFactory(Type itemType)
    {
        // Value types can always be created via Activator.CreateInstance
        if (itemType.IsValueType)
        {
            return () => Activator.CreateInstance(itemType);
        }

        // For reference types, prefer a compiled parameterless constructor if available
        var ctor = itemType.GetConstructor(Type.EmptyTypes);
        if (ctor is null)
        {
            // No default constructor - unable to materialize items automatically
            return () => null;
        }

        var newExpression = Expression.New(ctor);
        var lambda = Expression.Lambda<Func<object?>>(Expression.Convert(newExpression, typeof(object)));
        var factory = lambda.Compile();
        return factory;
    }

    /// <summary>
    /// Determines the item type from a collection type.
    /// </summary>
    /// <param name="collectionType">The collection type to analyze</param>
    /// <returns>The item type, or null if not determinable</returns>
    public static Type? GetItemType(Type collectionType)
    {
        if (collectionType.IsArray)
        {
            return collectionType.GetElementType();
        }

        if (!collectionType.IsGenericType) 
            return null;
        
        var genericArgs = collectionType.GetGenericArguments();
        return genericArgs.Length > 0 ? genericArgs[0] : null;
    }
}