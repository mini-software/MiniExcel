using System.ComponentModel;
using MiniExcelLib.Core.Attributes;
using MiniExcelLib.Core.Exceptions;

namespace MiniExcelLib.Core.Reflection;

internal static class ColumnMappingsProvider
{
    private const BindingFlags ExportMembersFlags = BindingFlags.Public | BindingFlags.Instance;
    private const BindingFlags ImportMembersFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty;


    internal static List<MiniExcelColumnMapping> GetMappingsForImport(Type type, string[] keys, MiniExcelBaseConfiguration configuration)
    {
        List<MiniExcelColumnMapping> mappings = GetColumnMappings(type, ImportMembersFlags, configuration)
            .Where(col => col?.MemberAccessor.CanWrite is true &&
                        !col.MemberAccessor.MemberInfo.GetAttributeValue((MiniExcelIgnoreAttribute x) => x.Ignore) &&
                        !col.MemberAccessor.MemberInfo.GetAttributeValue((MiniExcelColumnAttribute x) => x.Ignore))
            .ToList()!;

        if (mappings.Count == 0)
            throw new InvalidMappingException($"{type.Name} must contain at least one mappable property or field.", type);

        var firstDuplicateIndexGroup = mappings
            .Where(m => m?.ExcelColumnIndex > -1)
            .GroupBy(m => m?.ExcelColumnIndex)
            .FirstOrDefault(g => g.Count() > 1);

        if (firstDuplicateIndexGroup?.FirstOrDefault() is { } duplicate)
            throw new InvalidOperationException($"Duplicate column index in type {type.Name}: {duplicate.ExcelColumnIndex}");

        var maxKey = keys.Last();
        var maxIndex = CellReferenceConverter.GetNumericalIndex(maxKey);
        foreach (var p in mappings)
        {
            if (p?.ExcelColumnIndex is null)
                continue;

            if (p.ExcelColumnIndex > maxIndex)
                throw new InvalidMappingException($"The defined MiniExcelColumnIndex({p.ExcelColumnIndex}) exceeds the worksheets size({maxIndex})", type, p.MemberAccessor.MemberInfo);

            if (p.ExcelColumnName is null)
                throw new InvalidMappingException($"The defined MiniExcelColumnIndex({p.ExcelColumnIndex}) for type {type.Name}.{p.MemberAccessor.Name} does not match the defined MiniExcelColumnName({p.ExcelColumnName})", type, p.MemberAccessor.MemberInfo);
        }

        return mappings;
    }

    private static List<MiniExcelColumnMapping?> GetMappingsForExport(this Type type, MiniExcelBaseConfiguration configuration)
    {
        var props = GetColumnMappings(type, ExportMembersFlags, configuration)
            .Where(prop => prop?.MemberAccessor.CanRead is true)
            .ToList();

        if (props.Count == 0)
            throw new InvalidMappingException($"{type.Name} must contain at least one mappable property or field.", type);

        return SortMappings(props);
    }

    private static List<MiniExcelColumnMapping?> SortMappings(List<MiniExcelColumnMapping?> mappings)
    {
        //TODO: need optimize performance

        var explicitIndexMappings = mappings
            .Where(w => w?.ExcelColumnIndex > -1)
            .ToList();

        var firstDuplicateIndexGroup = mappings
            .Where(m => m?.ExcelColumnIndex > -1)
            .GroupBy(m => m?.ExcelColumnIndex)
            .FirstOrDefault(g => g.Count() > 1);

        if (firstDuplicateIndexGroup?.FirstOrDefault() is { } duplicate)
        {
            var type = duplicate.MemberAccessor.MemberInfo.DeclaringType;
            throw new InvalidMappingException($"Duplicate column index in type {type?.Name}: {duplicate.ExcelColumnIndex}", type, duplicate.MemberAccessor.MemberInfo);
        }

        var maxColumnIndex = mappings.Count - 1;
        if (explicitIndexMappings.Count != 0)
            maxColumnIndex = Math.Max(explicitIndexMappings.Max(w => w?.ExcelColumnIndex ?? 0), maxColumnIndex);

        var withoutCustomIndexProps = mappings
            .Where(w => w?.ExcelColumnIndex is null or -1)
            .ToList();

        var index = 0;
        var newProps = new List<MiniExcelColumnMapping?>();
        for (int i = 0; i <= maxColumnIndex; i++)
        {
            if (explicitIndexMappings.SingleOrDefault(s => s?.ExcelColumnIndex == i) is { } p1)
            {
                newProps.Add(p1);
            }
            else
            {
                var p2 = withoutCustomIndexProps.ElementAtOrDefault(index);

                p2?.ExcelColumnIndex = i;
                newProps.Add(p2);

                index++;
            }
        }
        return newProps;
    }

    private static IEnumerable<MiniExcelColumnMapping?> GetColumnMappings(Type type, BindingFlags bindingFlags, MiniExcelBaseConfiguration configuration)
    {
        var properties = type.GetProperties(bindingFlags);
        var fields = type.GetFields(bindingFlags).Where(x => x.GetCustomAttributes<MiniExcelAttributeBase>().Any());
        var members = properties.Cast<MemberInfo>().Concat(fields);
        
        var columnInfos = members.Select(m =>
        {
            var excelColumn = m.GetAttribute<MiniExcelColumnAttribute>();
            if (configuration.DynamicColumns?.SingleOrDefault(dc => dc.Key == m.Name) is { } dynamicColumn)
                excelColumn = dynamicColumn;

            var excelColumnName = m.GetAttribute<MiniExcelColumnNameAttribute>();
            var excelFormat = m.GetAttribute<MiniExcelFormatAttribute>()?.Format;

            var ignoreMember = 
                m.GetAttributeValue((MiniExcelIgnoreAttribute x) => x.Ignore) ||
                m.GetAttributeValue((MiniExcelColumnAttribute x) => x.Ignore) ||
                excelColumn?.Ignore is true;
            
            if (ignoreMember)
                return null;
                
            int? excelColumnIndex = excelColumn?.Index > -1 ? excelColumn.Index : null;
            var accessor = new MiniExcelMemberAccessor(m);
            
            //TODO: or dynamic configuration
            return new MiniExcelColumnMapping
            {
                MemberAccessor = new MiniExcelMemberAccessor(m),
                ExcludeNullableType = accessor.Type,
                Nullable = accessor.IsNullable,
                ExcelColumnAliases = excelColumnName?.Aliases ?? excelColumn?.Aliases ?? [],
                ExcelColumnName = excelColumnName?.ExcelColumnName ?? m.GetAttribute<DisplayNameAttribute>()?.DisplayName ?? excelColumn?.Name ?? m.Name,
                ExcelColumnIndex = m.GetAttribute<MiniExcelColumnIndexAttribute>()?.ExcelColumnIndex ?? excelColumnIndex,
                ExcelIndexName = m.GetAttribute<MiniExcelColumnIndexAttribute>()?.ExcelXName ?? excelColumn?.IndexName,
                ExcelColumnWidth = m.GetAttribute<MiniExcelColumnWidthAttribute>()?.ExcelColumnWidth ?? excelColumn?.Width,
                ExcelHiddenColumn = m.GetAttribute<MiniExcelHiddenAttribute>()?.Hidden ?? excelColumn?.Hidden ?? false,
                ExcelFormat = excelFormat ?? excelColumn?.Format,
                ExcelFormatId = excelColumn?.FormatId ?? -1,
                ExcelColumnType = excelColumn?.Type ?? ColumnType.Value,
                CustomFormatter = (excelColumn as DynamicExcelColumn)?.CustomFormatter
            };
        }); 
            
        return columnInfos.Where(x => x is not null);
    }

    private static List<MiniExcelColumnMapping?> GetDictionaryColumnInfo(IDictionary<string, object?>? dicString, IDictionary? dic, MiniExcelBaseConfiguration configuration)
    {
        var props = new List<MiniExcelColumnMapping?>();
        
        var keys = dicString?.Keys.ToList()
            ?? dic?.Keys
            ?? throw new InvalidOperationException();

        foreach (var key in keys)
        {
            SetDictionaryColumnInfo(props, key, configuration);
        }
        
        return SortMappings(props);
    }

    private static void SetDictionaryColumnInfo(List<MiniExcelColumnMapping?> props, object key, MiniExcelBaseConfiguration configuration)
    {
        var mapping = new MiniExcelColumnMapping
        {
            Key = key,
            ExcelColumnName = key?.ToString()
        };
            
        // TODO:Dictionary value type is not fixed
        var isIgnore = false;
        if (configuration.DynamicColumns is { Length: > 0 })
        {
            var dynamicColumn = configuration.DynamicColumns.SingleOrDefault(x => x.Key == key?.ToString());
            if (dynamicColumn is not null)
            {
                mapping.Nullable = true;

                if (dynamicColumn is { Format: { } fmt, FormatId: var fmtId })
                {
                    mapping.ExcelFormat = fmt;
                    mapping.ExcelFormatId = fmtId;
                }

                if (dynamicColumn.Aliases is { } aliases)
                    mapping.ExcelColumnAliases = aliases;

                if (dynamicColumn.IndexName is { } idxName)
                    mapping.ExcelIndexName = idxName;

                if (dynamicColumn.Name is { } colName)
                    mapping.ExcelColumnName = colName;

                mapping.ExcelColumnIndex = dynamicColumn.Index;
                mapping.ExcelColumnWidth = dynamicColumn.Width;
                mapping.ExcelHiddenColumn = dynamicColumn.Hidden;
                mapping.ExcelColumnType = dynamicColumn.Type;
                mapping.CustomFormatter = dynamicColumn.CustomFormatter;
                
                isIgnore = dynamicColumn.Ignore;
            }
        }
        
        if (!isIgnore)
            props.Add(mapping);
    }

    internal static bool TryGetColumnMappings(Type? type, MiniExcelBaseConfiguration configuration, out List<MiniExcelColumnMapping?> props)
    {
        props = [];

        // Unknown type
        if (type is null)
            return false;

        if (type.IsValueType || type == typeof(string))
            throw new NotSupportedException($"MiniExcel does not support the use of {type.FullName} as a generic type");

        if (ValueIsNeededToDetermineProperties(type))
            return false;

        props = type.GetMappingsForExport(configuration);
        return true;
    }
    
    internal static List<MiniExcelColumnMapping?> GetColumnMappingFromValue(object value, MiniExcelBaseConfiguration configuration) => value switch
    {
        IDictionary<string, object?> genericDictionary => GetDictionaryColumnInfo(genericDictionary, null, configuration),
        IDictionary dictionary => GetDictionaryColumnInfo(null, dictionary, configuration),
        _ => value.GetType().GetMappingsForExport(configuration)
    };

    private static bool ValueIsNeededToDetermineProperties(Type type) => 
        typeof(object) == type ||
        typeof(IDictionary<string, object>).IsAssignableFrom(type) ||
        typeof(IDictionary).IsAssignableFrom(type);

    internal static MiniExcelColumnMapping GetColumnMappingFromDynamicConfiguration(string columnName, MiniExcelBaseConfiguration configuration)
    {
        var member = new MiniExcelColumnMapping
        {
            ExcelColumnName = columnName,
            Key = columnName
        };

        if (configuration.DynamicColumns is null or [])
            return member;

        var dynamicColumn = configuration.DynamicColumns
            .SingleOrDefault(col => string.Equals(col.Key, columnName, StringComparison.OrdinalIgnoreCase));
            
        if (dynamicColumn is null)
            return member;

        member.Nullable = true;
        member.ExcelIgnoreColumn = dynamicColumn.Ignore;
        member.ExcelHiddenColumn = dynamicColumn.Hidden;
        member.ExcelColumnType = dynamicColumn.Type;
        member.ExcelColumnWidth = dynamicColumn.Width;
        member.CustomFormatter = dynamicColumn.CustomFormatter;

        if (dynamicColumn is { Format: { } fmt, FormatId: var fmtId })
        {
            member.ExcelFormat = fmt;
            member.ExcelFormatId = fmtId;
        }

        if (dynamicColumn.Index > -1)
            member.ExcelColumnIndex = dynamicColumn.Index;

        if (dynamicColumn.Aliases is { } aliases)
            member.ExcelColumnAliases = aliases;

        if (dynamicColumn.IndexName is { } idxName)
            member.ExcelIndexName = idxName;

        if (dynamicColumn.Name is { } colName)
            member.ExcelColumnName = colName;

        return member;
    }
        
    internal static Dictionary<string, int> GetHeaders(IDictionary<string, object?> item, bool trimNames = false)
    {
        return DictToNameWithIndex(item)
            .GroupBy(x => x.Name)
            .SelectMany(GroupToNameWithIndex)
            .ToDictionary(kv => trimNames ? kv.Name.Trim() : kv.Name, kv => kv.Index);
        
        static IEnumerable<NameIndexPair> DictToNameWithIndex(IDictionary<string, object?> dict)
            => dict.Values.Select((obj, idx) => new NameIndexPair(idx, obj?.ToString() ?? ""));
        
        static IEnumerable<NameIndexPair> GroupToNameWithIndex(IGrouping<string, NameIndexPair> group)
            => group.Select((grp, idx) => new NameIndexPair(grp.Index, idx == 0 ? grp.Name : $"{grp.Name}_____{idx + 1}"));
    }
        
    private class NameIndexPair(int index, string name)
    {
        public int Index { get; } = index;
        public string Name { get; } = name;
    }
}
