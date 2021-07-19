namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.Attributes;
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Linq;
    using System.Reflection;

    internal class ExcelCustomPropertyInfo
    {
        public int? ExcelColumnIndex { get; set; }
        public string ExcelColumnName { get; set; }
        public PropertyInfo Property { get; set; }
        public Type ExcludeNullableType { get; set; }
        public bool Nullable { get; internal set; }
        public string ExcelFormat { get; internal set; }
    }

    internal static partial class CustomPropertyHelper
    {
        internal static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex, int startCellIndex)
        {
            // TODO: strong type mapping can ignore this
            // TODO: it can recode better performance 
            var cell = (IDictionary<string, object>)new ExpandoObject();
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
            // TODO: strong type mapping can ignore this
            // TODO: it can recode better performance 
            var cell = (IDictionary<string, object>)new ExpandoObject();
            foreach (var hr in hearrows)
                if (!cell.ContainsKey(hr.Value))
                    cell.Add(hr.Value, null);
            return cell;
        }

        internal static List<ExcelCustomPropertyInfo> GetSaveAsProperties(this Type type)
        {
            List<ExcelCustomPropertyInfo> props = GetExcelPropertyInfo(type, BindingFlags.Public | BindingFlags.Instance)
                .Where(prop => prop.Property.GetGetMethod() != null && !prop.Property.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore))
                .ToList() /*ignore without set*/;

            if (props.Count == 0)
                throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

            // https://github.com/shps951023/MiniExcel/issues/142
            //TODO: need optimize performance
            {
                var withCustomIndexProps = props.Where(w => w.ExcelColumnIndex != null && w.ExcelColumnIndex > -1);
                if (withCustomIndexProps.GroupBy(g => g.ExcelColumnIndex).Any(_ => _.Count() > 1))
                    throw new InvalidOperationException($"Duplicate column name");

                var maxColumnIndex = props.Count - 1;
                if (withCustomIndexProps.Any())
                    maxColumnIndex = Math.Max((int)withCustomIndexProps.Max(w => w.ExcelColumnIndex), maxColumnIndex);

                var withoutCustomIndexProps = props.Where(w => w.ExcelColumnIndex == null).ToList();

                List<ExcelCustomPropertyInfo> newProps = new List<ExcelCustomPropertyInfo>();
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
        }

        internal static List<ExcelCustomPropertyInfo> GetExcelCustomPropertyInfos(Type type, string[] headers)
        {
            List<ExcelCustomPropertyInfo> props = GetExcelPropertyInfo(type, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance)
                .Where(prop => prop.Property.GetSetMethod() != null && !prop.Property.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore))
                .ToList() /*ignore without set*/;

            if (props.Count == 0)
                throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

            {
                var withCustomIndexProps = props.Where(w => w.ExcelColumnIndex != null && w.ExcelColumnIndex > -1);
                if (withCustomIndexProps.GroupBy(g => g.ExcelColumnIndex).Any(_ => _.Count() > 1))
                    throw new InvalidOperationException($"Duplicate column name");

                foreach (var p in props)
                {
                    if (p.ExcelColumnIndex != null)
                    {
                        if (p.ExcelColumnIndex >= headers.Length)
                            throw new ArgumentException($"ExcelColumnIndex {p.ExcelColumnIndex} over haeder max index {headers.Length}");
                        p.ExcelColumnName = headers[(int)p.ExcelColumnIndex];
                        if (p.ExcelColumnName == null)
                            throw new InvalidOperationException($"{p.Property.DeclaringType.Name} {p.Property.Name}'s ExcelColumnIndex {p.ExcelColumnIndex} can't find excel column name");
                    }
                }
            }

            return props;
        }

        private static IEnumerable<ExcelCustomPropertyInfo> GetExcelPropertyInfo(Type type, BindingFlags bindingFlags)
        {
            return type.GetProperties(bindingFlags)
                 // solve : https://github.com/shps951023/MiniExcel/issues/138
                 .Select(p =>
                 {
                     var gt = Nullable.GetUnderlyingType(p.PropertyType);
                     var excelNameAttr = p.GetAttribute<ExcelColumnNameAttribute>();
                     var excelIndexAttr = p.GetAttribute<ExcelColumnIndexAttribute>();
                     return new ExcelCustomPropertyInfo
                     {
                         Property = p,
                         ExcludeNullableType = gt ?? p.PropertyType,
                         Nullable = gt != null ? true : false,
                         ExcelColumnName = excelNameAttr?.ExcelColumnName ?? p.Name,
                         ExcelColumnIndex = excelIndexAttr?.ExcelColumnIndex,
                         ExcelFormat = p.GetAttribute<ExcelFormatAttribute>()?.Format,
                     };
                 });
        }

    }

}
