namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.Attributes;
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.OpenXml.Models;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;

    internal class ExcelColumnInfo
    {
        public object Key { get; set; }
        public int? ExcelColumnIndex { get; set; }
        public string ExcelColumnName { get; set; }
        public string[] ExcelColumnAliases { get; set; }
        public Property Property { get; set; }
        public Type ExcludeNullableType { get; set; }
        public bool Nullable { get; internal set; }
        public string ExcelFormat { get; internal set; }
        public double? ExcelColumnWidth { get; internal set; }
        public string ExcelIndexName { get; internal set; }
        public bool ExcelIgnore { get; internal set; }
    }

    internal class ExcellSheetInfo
    {
        public object Key { get; set; }
        public string ExcelSheetName { get; set; }
        public SheetState ExcelSheetState { get; set; }

        private string ExcelSheetStateAsString
        {
            get
            {
                return ExcelSheetState.ToString().ToLower();
            }
        }

        public SheetDto ToDto(int sheetIndex)
        {
            return new SheetDto { Name = ExcelSheetName, SheetIdx = sheetIndex, State = ExcelSheetStateAsString };
        }
    }

    internal static partial class CustomPropertyHelper
    {
        internal static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex, int startCellIndex)
        {
            var cell = new Dictionary<string, object>();
            for (int i = startCellIndex; i <= maxColumnIndex; i++)
            {
                var key = ColumnHelper.GetAlphabetColumnName(i);
                if (!cell.ContainsKey(key))
                    cell.Add(key, null);
            }
            return cell;
        }

        internal static IDictionary<string, object> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
        {
            var cell = new Dictionary<string, object>();
            foreach (var hr in hearrows)
                if (!cell.ContainsKey(hr.Value))
                    cell.Add(hr.Value, null);
            return cell;
        }

        internal static List<ExcelColumnInfo> GetSaveAsProperties(this Type type, Configuration configuration)
        {
            List<ExcelColumnInfo> props = GetExcelPropertyInfo(type, BindingFlags.Public | BindingFlags.Instance, configuration)
                .Where(prop => prop.Property.CanRead)
                .ToList() /*ignore without set*/;

            if (props.Count == 0)
                throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

            return SortCustomProps(props);
        }

        internal static List<ExcelColumnInfo> SortCustomProps(List<ExcelColumnInfo> props)
        {
            // https://github.com/shps951023/MiniExcel/issues/142
            //TODO: need optimize performance

            var withCustomIndexProps = props.Where(w => w.ExcelColumnIndex != null && w.ExcelColumnIndex > -1);
            if (withCustomIndexProps.GroupBy(g => g.ExcelColumnIndex).Any(_ => _.Count() > 1))
                throw new InvalidOperationException($"Duplicate column name");

            var maxColumnIndex = props.Count - 1;
            if (withCustomIndexProps.Any())
                maxColumnIndex = Math.Max((int)withCustomIndexProps.Max(w => w.ExcelColumnIndex), maxColumnIndex);

            var withoutCustomIndexProps = props.Where(w => w.ExcelColumnIndex == null || w.ExcelColumnIndex == -1).ToList();

            List<ExcelColumnInfo> newProps = new List<ExcelColumnInfo>();
            var index = 0;
            for (int i = 0; i <= maxColumnIndex; i++)
            {
                var p1 = withCustomIndexProps.SingleOrDefault(s => s.ExcelColumnIndex == i);
                if (p1 != null)
                {
                    newProps.Add(p1);
                }
                else
                {
                    var p2 = withoutCustomIndexProps.ElementAtOrDefault(index);
                    if (p2 == null)
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

        internal static List<ExcelColumnInfo> GetExcelCustomPropertyInfos(Type type, string[] keys, Configuration configuration)
        {
            List<ExcelColumnInfo> props = GetExcelPropertyInfo(type, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance, configuration)
                .Where(prop => prop.Property.Info.GetSetMethod() != null // why not .Property.CanWrite? because it will use private setter
                               && !prop.Property.Info.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore)
                               && !prop.Property.Info.GetAttributeValue((ExcelColumnAttribute x) => x.Ignore))
                .ToList() /*ignore without set*/;

            if (props.Count == 0)
                throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

            {
                var withCustomIndexProps = props.Where(w => w.ExcelColumnIndex != null && w.ExcelColumnIndex > -1);
                if (withCustomIndexProps.GroupBy(g => g.ExcelColumnIndex).Any(_ => _.Count() > 1))
                    throw new InvalidOperationException($"Duplicate column name");
                var maxkey = keys.Last();
                var maxIndex = ColumnHelper.GetColumnIndex(maxkey);
                foreach (var p in props)
                {
                    if (p.ExcelColumnIndex != null)
                    {
                        if (p.ExcelColumnIndex > maxIndex)
                            throw new ArgumentException($"ExcelColumnIndex {p.ExcelColumnIndex} over haeder max index {maxkey}");
                        if (p.ExcelColumnName == null)
                            throw new InvalidOperationException($"{p.Property.Info.DeclaringType.Name} {p.Property.Name}'s ExcelColumnIndex {p.ExcelColumnIndex} can't find excel column name");
                    }
                }
            }

            return props;
        }

        internal static string DescriptionAttr(Type type, object source)
        {
            FieldInfo fi = type.GetField(source.ToString());
            //For some database dirty data, there may be no way to change to the correct enumeration, will return NULL
            if (fi == null)
                return source.ToString();

            DescriptionAttribute[] attributes = (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return source.ToString();
        }

        private static IEnumerable<ExcelColumnInfo> ConvertToExcelCustomPropertyInfo(PropertyInfo[] props, Configuration configuration)
        {
            // solve : https://github.com/shps951023/MiniExcel/issues/138
            return props.Select(p =>
            {
                var gt = Nullable.GetUnderlyingType(p.PropertyType);
                var excelColumnName = p.GetAttribute<ExcelColumnNameAttribute>();
                var excludeNullableType = gt ?? p.PropertyType;
                var excelFormat = p.GetAttribute<ExcelFormatAttribute>()?.Format;
                var excelColumn = p.GetAttribute<ExcelColumnAttribute>();
                if (configuration.DynamicColumns != null && configuration.DynamicColumns.Length > 0)
                {
                    var dynamicColumn = configuration.DynamicColumns.SingleOrDefault(_ => _.Key == p.Name);
                    if (dynamicColumn != null)
                        excelColumn = dynamicColumn;
                }

                var ignore = p.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore) || p.GetAttributeValue((ExcelColumnAttribute x) => x.Ignore) || (excelColumn != null && excelColumn.Ignore);
                if (ignore)
                {
                    return null;
                }
                //TODO:or configulation Dynamic
                var excelColumnIndex = excelColumn?.Index > -1 ? excelColumn.Index : (int?)null;
                return new ExcelColumnInfo
                {
                    Property = new Property(p),
                    ExcludeNullableType = excludeNullableType,
                    Nullable = gt != null,
                    ExcelColumnAliases = excelColumnName?.Aliases ?? excelColumn?.Aliases,
                    ExcelColumnName = excelColumnName?.ExcelColumnName ?? p.GetAttribute<DisplayNameAttribute>()?.DisplayName ?? excelColumn?.Name ?? p.Name,
                    ExcelColumnIndex = p.GetAttribute<ExcelColumnIndexAttribute>()?.ExcelColumnIndex ?? excelColumnIndex,
                    ExcelIndexName = p.GetAttribute<ExcelColumnIndexAttribute>()?.ExcelXName ?? excelColumn?.IndexName,
                    ExcelColumnWidth = p.GetAttribute<ExcelColumnWidthAttribute>()?.ExcelColumnWidth ?? excelColumn?.Width,
                    ExcelFormat = excelFormat ?? excelColumn?.Format,
                };
            }).Where(_ => _ != null);
        }

        private static IEnumerable<ExcelColumnInfo> GetExcelPropertyInfo(Type type, BindingFlags bindingFlags, Configuration configuration)
        {
            //TODO:assign column index
            return ConvertToExcelCustomPropertyInfo(type.GetProperties(bindingFlags), configuration);
        }

        internal static ExcellSheetInfo GetExcellSheetInfo(Type type, Configuration configuration)
        {
            // default options
            var sheetInfo = new ExcellSheetInfo()
            {
                Key = type.Name,
                ExcelSheetName = null, // will be generated automatically as Sheet<Index>
                ExcelSheetState = SheetState.Visible
            };

            // options from ExcelSheetAttribute
            ExcelSheetAttribute excelSheetAttribute = type.GetCustomAttribute(typeof(ExcelSheetAttribute)) as ExcelSheetAttribute;
            if (excelSheetAttribute != null)
            {
                sheetInfo.ExcelSheetName = excelSheetAttribute.Name ?? type.Name;
                sheetInfo.ExcelSheetState = excelSheetAttribute.State;
            }

            // options from DynamicSheets configuration
            OpenXmlConfiguration openXmlCOnfiguration = configuration as OpenXmlConfiguration;
            if (openXmlCOnfiguration != null && openXmlCOnfiguration.DynamicSheets != null && openXmlCOnfiguration.DynamicSheets.Length > 0)
            {
                var dynamicSheet = openXmlCOnfiguration.DynamicSheets.SingleOrDefault(_ => _.Key == type.Name);
                if (dynamicSheet != null)
                {
                    sheetInfo.ExcelSheetName = dynamicSheet.Name;
                    sheetInfo.ExcelSheetState = dynamicSheet.State;
                }
            }

            return sheetInfo;
        }

        internal static List<ExcelColumnInfo> GetDictionaryColumnInfo(IDictionary<string, object> dicString, IDictionary dic, Configuration configuration)
        {
            List<ExcelColumnInfo> props;
            var _props = new List<ExcelColumnInfo>();
            if (dicString != null)
            {
                foreach (var key in dicString.Keys)
                {
                    SetDictionaryColumnInfo(_props, key, configuration);
                }
            }
            else if (dic != null)
            {
                foreach (var key in dic.Keys)
                {
                    SetDictionaryColumnInfo(_props, key, configuration);
                }
            }
            else
            {
                throw new NotSupportedException("SetDictionaryColumnInfo Error");
            }

            props = SortCustomProps(_props);
            return props;
        }

        internal static void SetDictionaryColumnInfo(List<ExcelColumnInfo> _props, object key, Configuration configuration)
        {
            var p = new ExcelColumnInfo();
            p.ExcelColumnName = key?.ToString();
            p.Key = key;
            // TODO:Dictionary value type is not fiexed
            //var _t =
            //var gt = Nullable.GetUnderlyingType(p.PropertyType);
            var isIgnore = false;
            if (configuration.DynamicColumns != null && configuration.DynamicColumns.Length > 0)
            {
                var dynamicColumn = configuration.DynamicColumns.SingleOrDefault(_ => _.Key == key.ToString());
                if (dynamicColumn != null)
                {
                    p.Nullable = true;
                    //p.ExcludeNullableType = item2[key]?.GetType();
                    if (dynamicColumn.Format != null)
                        p.ExcelFormat = dynamicColumn.Format;
                    if (dynamicColumn.Aliases != null)
                        p.ExcelColumnAliases = dynamicColumn.Aliases;
                    if (dynamicColumn.IndexName != null)
                        p.ExcelIndexName = dynamicColumn.IndexName;
                    p.ExcelColumnIndex = dynamicColumn.Index;
                    if (dynamicColumn.Name != null)
                        p.ExcelColumnName = dynamicColumn.Name;
                    isIgnore = dynamicColumn.Ignore;
                    p.ExcelColumnWidth = dynamicColumn.Width;
                }
            }
            if (!isIgnore)
                _props.Add(p);
        }
    }
}