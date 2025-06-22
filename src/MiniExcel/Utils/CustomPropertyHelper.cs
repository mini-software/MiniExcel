﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.OpenXml.Models;

namespace MiniExcelLibs.Utils;

internal class ExcelColumnInfo
{
    public object Key { get; set; }
    public int? ExcelColumnIndex { get; set; }
    public string? ExcelColumnName { get; set; }
    public string[]? ExcelColumnAliases { get; set; } = [];
    public Property Property { get; set; }
    public Type ExcludeNullableType { get; set; }
    public bool Nullable { get; internal set; }
    public string? ExcelFormat { get; internal set; }
    public double? ExcelColumnWidth { get; internal set; }
    public string? ExcelIndexName { get; internal set; }
    public bool ExcelIgnore { get; internal set; }
    public int ExcelFormatId { get; internal set; }
    public ColumnType ExcelColumnType { get; internal set; }
    public Func<object, object>? CustomFormatter { get; set; }
}

internal class ExcellSheetInfo
{
    public object Key { get; set; }
    public string? ExcelSheetName { get; set; }
    public SheetState ExcelSheetState { get; set; }
        
    private string ExcelSheetStateAsString => ExcelSheetState.ToString().ToLower();

    public SheetDto ToDto(int sheetIndex)
    {
        return new SheetDto { Name = ExcelSheetName, SheetIdx = sheetIndex, State = ExcelSheetStateAsString };
    }
}

internal static class CustomPropertyHelper
{
    internal static IDictionary<string, object?> GetEmptyExpandoObject(int maxColumnIndex, int startCellIndex)
    {
        var cell = new Dictionary<string, object?>();
        for (int i = startCellIndex; i <= maxColumnIndex; i++)
        {
            var key = ColumnHelper.GetAlphabetColumnName(i);
#if NETCOREAPP2_0_OR_GREATER
            cell.TryAdd(key, null);
#else
            if (!cell.ContainsKey(key))
                cell.Add(key, null);
#endif
        }
        return cell;
    }

    internal static IDictionary<string, object?> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
    {
        var cell = new Dictionary<string, object?>();
        foreach (var hr in hearrows)
        {
#if NETCOREAPP2_0_OR_GREATER
            cell.TryAdd(hr.Value, null);
#else
            if (!cell.ContainsKey(hr.Value))
                cell.Add(hr.Value, null);
#endif
        }

        return cell;
    }

    internal static List<ExcelColumnInfo> GetSaveAsProperties(this Type type, MiniExcelConfiguration configuration)
    {
        var props = GetExcelPropertyInfo(type, BindingFlags.Public | BindingFlags.Instance, configuration)
            .Where(prop => prop.Property.CanRead)
            .ToList() /*ignore without set*/;

        if (props.Count == 0)
            throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

        return SortCustomProps(props);
    }

    internal static List<ExcelColumnInfo?> SortCustomProps(List<ExcelColumnInfo> props)
    {
        // https://github.com/mini-software/MiniExcel/issues/142
        //TODO: need optimize performance

        var withCustomIndexProps = props.Where(w => w.ExcelColumnIndex is > -1).ToList();
        if (withCustomIndexProps.GroupBy(g => g.ExcelColumnIndex).Any(x => x.Count() > 1))
            throw new InvalidOperationException("Duplicate column name");

        var maxColumnIndex = props.Count - 1;
        if (withCustomIndexProps.Count != 0)
            maxColumnIndex = Math.Max((int)withCustomIndexProps.Max(w => w.ExcelColumnIndex), maxColumnIndex);

        var withoutCustomIndexProps = props.Where(w => w.ExcelColumnIndex is null or -1).ToList();

        var index = 0;
        var newProps = new List<ExcelColumnInfo?>();
        for (int i = 0; i <= maxColumnIndex; i++)
        {
            var p1 = withCustomIndexProps.SingleOrDefault(s => s.ExcelColumnIndex == i);
            if (p1 is not null)
            {
                newProps.Add(p1);
            }
            else
            {
                var p2 = withoutCustomIndexProps.ElementAtOrDefault(index);
                if (p2 is null)
                {
                    newProps.Add(null);
                }
                else
                {
                    p2.ExcelColumnIndex = i;
                    newProps.Add(p2);
                }
                index++;
            }
        }
        return newProps;
    }

    internal static List<ExcelColumnInfo?> GetExcelCustomPropertyInfos(Type type, string[] keys, MiniExcelConfiguration configuration)
    {
        const BindingFlags flags = BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance;
        var props = GetExcelPropertyInfo(type, flags, configuration)
            .Where(prop => prop?.Property.Info.GetSetMethod() is not null // why not .Property.CanWrite? because it will use private setter
                           && !prop.Property.Info.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore)
                           && !prop.Property.Info.GetAttributeValue((ExcelColumnAttribute x) => x.Ignore))
            .ToList() /*ignore without set*/;

        if (props.Count == 0)
            throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

        var withCustomIndexProps = props.Where(w => w?.ExcelColumnIndex is > -1);
        if (withCustomIndexProps.GroupBy(g => g?.ExcelColumnIndex).Any(x => x.Count() > 1))
            throw new InvalidOperationException("Duplicate column name");
            
        var maxkey = keys.Last();
        var maxIndex = ColumnHelper.GetColumnIndex(maxkey);
        foreach (var p in props)
        {
            if (p?.ExcelColumnIndex is null)
                continue;
            if (p.ExcelColumnIndex > maxIndex)
                throw new ArgumentException($"ExcelColumnIndex {p.ExcelColumnIndex} over haeder max index {maxkey}");
            if (p.ExcelColumnName is null)
                throw new InvalidOperationException($"{p.Property.Info.DeclaringType?.Name} {p.Property.Name}'s ExcelColumnIndex {p.ExcelColumnIndex} can't find excel column name");
        }

        return props;
    }

    internal static string? DescriptionAttr(Type type, object? source)
    {
        var name = source?.ToString();
        return type.GetField(name) //For some database dirty data, there may be no way to change to the correct enumeration, will return NULL
           ?.GetCustomAttribute<DescriptionAttribute>(false)?.Description 
               ?? name;
    }

    private static IEnumerable<ExcelColumnInfo?> ConvertToExcelCustomPropertyInfo(PropertyInfo[] props, MiniExcelConfiguration configuration)
    {
        // solve : https://github.com/mini-software/MiniExcel/issues/138
        var columnInfos = props.Select(p =>
        {
            var gt = Nullable.GetUnderlyingType(p.PropertyType);
            var excelColumnName = p.GetAttribute<ExcelColumnNameAttribute>();
            var excludeNullableType = gt ?? p.PropertyType;
            var excelFormat = p.GetAttribute<ExcelFormatAttribute>()?.Format;
            var excelColumn = p.GetAttribute<ExcelColumnAttribute>();
            var dynamicColumn = configuration?.DynamicColumns?.SingleOrDefault(_ => _.Key == p.Name);
            if (dynamicColumn is not null)
                excelColumn = dynamicColumn;

            var ignore = p.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore) ||
                         p.GetAttributeValue((ExcelColumnAttribute x) => x.Ignore) ||
                         (excelColumn?.Ignore ?? false);
            if (ignore)
                return null;
                
            //TODO:or configulation Dynamic
            int? excelColumnIndex = excelColumn?.Index > -1 ? excelColumn.Index : null;
            return new ExcelColumnInfo
            {
                Property = new Property(p),
                ExcludeNullableType = excludeNullableType,
                Nullable = gt is not null,
                ExcelColumnAliases = excelColumnName?.Aliases ?? excelColumn?.Aliases ?? [],
                ExcelColumnName = excelColumnName?.ExcelColumnName ?? p.GetAttribute<DisplayNameAttribute>()?.DisplayName ?? excelColumn?.Name ?? p.Name,
                ExcelColumnIndex = p.GetAttribute<ExcelColumnIndexAttribute>()?.ExcelColumnIndex ?? excelColumnIndex,
                ExcelIndexName = p.GetAttribute<ExcelColumnIndexAttribute>()?.ExcelXName ?? excelColumn?.IndexName,
                ExcelColumnWidth = p.GetAttribute<ExcelColumnWidthAttribute>()?.ExcelColumnWidth ?? excelColumn?.Width,
                ExcelFormat = excelFormat ?? excelColumn?.Format,
                ExcelFormatId = excelColumn?.FormatId ?? -1,
                ExcelColumnType = excelColumn?.Type ?? ColumnType.Value,
                CustomFormatter = dynamicColumn?.CustomFormatter
            };
        }); 
            
        return columnInfos.Where(x => x is not null);
    }

    private static IEnumerable<ExcelColumnInfo?> GetExcelPropertyInfo(Type type, BindingFlags bindingFlags, MiniExcelConfiguration configuration)
    {
        //TODO:assign column index
        return ConvertToExcelCustomPropertyInfo(type.GetProperties(bindingFlags), configuration);
    }

    internal static ExcellSheetInfo GetExcellSheetInfo(Type type, MiniExcelConfiguration configuration)
    {
        // default options
        var sheetInfo = new ExcellSheetInfo
        {
            Key = type.Name,
            ExcelSheetName = null, // will be generated automatically as Sheet<Index>
            ExcelSheetState = SheetState.Visible
        };

        // options from ExcelSheetAttribute
        if (type.GetCustomAttribute(typeof(ExcelSheetAttribute)) is ExcelSheetAttribute excelSheetAttr)
        {
            sheetInfo.ExcelSheetName = excelSheetAttr.Name ?? type.Name;
            sheetInfo.ExcelSheetState = excelSheetAttr.State;
        }

        // options from DynamicSheets configuration
        var openXmlCOnfiguration = configuration as OpenXmlConfiguration;
        if (openXmlCOnfiguration?.DynamicSheets?.Length > 0)
        {
            var dynamicSheet = openXmlCOnfiguration.DynamicSheets.SingleOrDefault(x => x.Key == type.Name);
            if (dynamicSheet is not null)
            {
                sheetInfo.ExcelSheetName = dynamicSheet.Name;
                sheetInfo.ExcelSheetState = dynamicSheet.State;
            }
        }

        return sheetInfo;
    }

    internal static List<ExcelColumnInfo> GetDictionaryColumnInfo(IDictionary<string, object?>? dicString, IDictionary? dic, MiniExcelConfiguration configuration)
    {
        var props = new List<ExcelColumnInfo>();
            
        var keys = dicString?.Keys.ToList() 
                   ?? dic?.Keys
                   ?? throw new NotSupportedException();

        foreach (var key in keys)
        {
            SetDictionaryColumnInfo(props, key, configuration);
        }
        
        return SortCustomProps(props);
    }

    internal static void SetDictionaryColumnInfo(List<ExcelColumnInfo> props, object key, MiniExcelConfiguration configuration)
    {
        var p = new ExcelColumnInfo
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
                p.Nullable = true;

                if (dynamicColumn.Format is not null)
                {
                    p.ExcelFormat = dynamicColumn.Format;
                    p.ExcelFormatId = dynamicColumn.FormatId;
                }
                
                if (dynamicColumn.Aliases is not null)
                    p.ExcelColumnAliases = dynamicColumn.Aliases;
                
                if (dynamicColumn.IndexName is not null)
                    p.ExcelIndexName = dynamicColumn.IndexName;
                
                if (dynamicColumn.Name is not null)
                    p.ExcelColumnName = dynamicColumn.Name;
                
                p.ExcelColumnIndex = dynamicColumn.Index;
                p.ExcelColumnWidth = dynamicColumn.Width;
                p.ExcelColumnType = dynamicColumn.Type;
                p.CustomFormatter = dynamicColumn.CustomFormatter;
                
                isIgnore = dynamicColumn.Ignore;
            }
        }
        
        if (!isIgnore)
            props.Add(p);
    }

    internal static bool TryGetTypeColumnInfo(Type? type, MiniExcelConfiguration configuration, out List<ExcelColumnInfo>? props)
    {
        // Unknown type
        if (type is null)
        {
            props = null;
            return false;
        }

        if (type.IsValueType || type == typeof(string))
            throw new NotSupportedException($"MiniExcel does not support the use of {type.FullName} as a generic type");

        if (ValueIsNeededToDetermineProperties(type))
        {
            props = null;
            return false;
        }

        props = GetSaveAsProperties(type, configuration);
        return true;
    }
    internal static List<ExcelColumnInfo> GetColumnInfoFromValue(object value, MiniExcelConfiguration configuration) => value switch
    {
        IDictionary<string, object?> genericDictionary => GetDictionaryColumnInfo(genericDictionary, null, configuration),
        IDictionary dictionary => GetDictionaryColumnInfo(null, dictionary, configuration),
        _ => GetSaveAsProperties(value.GetType(), configuration)
    };

    private static bool ValueIsNeededToDetermineProperties(Type type) => 
        typeof(object) == type ||
        typeof(IDictionary<string, object>).IsAssignableFrom(type) ||
        typeof(IDictionary).IsAssignableFrom(type);

    internal static ExcelColumnInfo GetColumnInfosFromDynamicConfiguration(string columnName, MiniExcelConfiguration configuration)
    {
        var prop = new ExcelColumnInfo
        {
            ExcelColumnName = columnName,
            Key = columnName
        };

        if (configuration.DynamicColumns is null or [])
            return prop;

        var dynamicColumn = configuration.DynamicColumns
            .SingleOrDefault(col => string.Equals(col.Key, columnName, StringComparison.OrdinalIgnoreCase));
            
        if (dynamicColumn is null)
            return prop;

        prop.Nullable = true;
        prop.ExcelIgnore = dynamicColumn.Ignore;
        prop.ExcelColumnType = dynamicColumn.Type;
        prop.ExcelColumnWidth = dynamicColumn.Width;
        prop.CustomFormatter = dynamicColumn.CustomFormatter;

        if (dynamicColumn.Index > -1)
        {
            prop.ExcelColumnIndex = dynamicColumn.Index;
        }
            
        if (dynamicColumn.Format is not null)
        {
            prop.ExcelFormat = dynamicColumn.Format;
            prop.ExcelFormatId = dynamicColumn.FormatId;
        }

        if (dynamicColumn.Aliases is not null)
        {
            prop.ExcelColumnAliases = dynamicColumn.Aliases;
        }

        if (dynamicColumn.IndexName is not null)
        {
            prop.ExcelIndexName = dynamicColumn.IndexName;
        }

        if (dynamicColumn.Name is not null)
        {
            prop.ExcelColumnName = dynamicColumn.Name;
        }

        return prop;
    }
        
    internal static Dictionary<string, int> GetHeaders(IDictionary<string, object?> item, bool trimNames = false)
    {
        return DictToNameWithIndex(item)
            .GroupBy(x => x.Name)
            .SelectMany(GroupToNameWithIndex)
            .ToDictionary(kv => trimNames ? kv.Name.Trim() : kv.Name, kv => kv.Index);
    }

    private static IEnumerable<NameIndexPair> DictToNameWithIndex(IDictionary<string, object?> dict)
    {
        return dict.Values.Select((obj, idx) => new NameIndexPair(idx, obj?.ToString() ?? ""));
    }
        
    private static IEnumerable<NameIndexPair> GroupToNameWithIndex(IGrouping<string, NameIndexPair> group)
    {
        return group.Select((grp, idx) =>
            new NameIndexPair(grp.Index, idx == 0 ? grp.Name : $"{grp.Name}_____{idx + 1}"));
    }
        
    private class NameIndexPair(int index, string name)
    {
        public int Index { get; } = index;
        public string Name { get; } = name;
    }
}