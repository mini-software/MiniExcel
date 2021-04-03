namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.Attributes;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;
    using System.Text.RegularExpressions;

    internal static partial class Helpers
    {
        private static Dictionary<int, string> _IntMappingAlphabet = new Dictionary<int, string>();
        private static Dictionary<string, int> _AlphabetMappingInt = new Dictionary<string, int>();
        static Helpers()
        {
            for (int i = 0; i <= 255; i++)
            {
                _IntMappingAlphabet.Add(i, IntToLetters(i));
                _AlphabetMappingInt.Add(IntToLetters(i), i);
            }
        }

        public static string GetAlphabetColumnName(int ColumnIndex) => _IntMappingAlphabet[ColumnIndex];
        public static int GetColumnIndex(string columnName) => _AlphabetMappingInt[columnName];

        internal static string IntToLetters(int value)
        {
            value = value + 1;
            string result = string.Empty;
            while (--value >= 0)
            {
                result = (char)('A' + value % 26) + result;
                value /= 26;
            }
            return result;
        }

        public static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex)
        {
            // TODO: strong type mapping can ignore this
            // TODO: it can recode better performance 
            var cell = (IDictionary<string, object>)new ExpandoObject();
            for (int i = 0; i <= maxColumnIndex; i++)
            {
                var key = GetAlphabetColumnName(i);
                if (!cell.ContainsKey(key))
                    cell.Add(key, null);
            }
            return cell;
        }

        public static IDictionary<string, object> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
        {
            // TODO: strong type mapping can ignore this
            // TODO: it can recode better performance 
            var cell = (IDictionary<string, object>)new ExpandoObject();
            foreach (var hr in hearrows)
                if (!cell.ContainsKey(hr.Value))
                    cell.Add(hr.Value, null);
            return cell;
        }

        public static PropertyInfo[] GetProperties(this Type type)
        {
            return type.GetProperties(
                         BindingFlags.Public |
                         BindingFlags.Instance);
        }

        internal class ExcelCustomPropertyInfo
        {
            public int? ExcelColumnIndex { get; set; }
            public string ExcelColumnName { get; set; }
            public PropertyInfo Property { get; set; }
            public Type ExcludeNullableType { get; set; }
            public bool Nullable { get; internal set; }
        }

        internal static List<ExcelCustomPropertyInfo> GetExcelCustomPropertyInfos(Type type)
        {
            List<ExcelCustomPropertyInfo> props = GetExcelPropertyInfo(type, BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance)
                .Where(prop => prop.Property.IsSupportSetMethod() && !prop.Property.GetAttributeValue((ExcelIgnoreAttribute x) => x.ExcelIgnore))
                .ToList() /*ignore without set*/;

            if (props.Count == 0)
                throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

            return props;
        }

        private static IEnumerable<ExcelCustomPropertyInfo> GetExcelPropertyInfo(Type type, BindingFlags bindingFlags)
        {
            return type.GetProperties(bindingFlags)
                 // solve : https://github.com/shps951023/MiniExcel/issues/138
                 .Select(p =>
                 {
                     var gt = Nullable.GetUnderlyingType(p.PropertyType);
                     var excelAttr = p.GetAttribute<ExcelColumnNameAttribute>();
                     return new ExcelCustomPropertyInfo
                     {
                         Property = p,
                         ExcludeNullableType = gt ?? p.PropertyType,
                         Nullable = gt != null ? true : false,
                         ExcelColumnName = excelAttr?.ExcelColumnName ?? p.Name
                     };
                 });
        }

        internal static bool IsDapperRows<T>()
        {
            return typeof(IDictionary<string, object>).IsAssignableFrom(typeof(T));
        }

        private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");
        public static string ConvertEscapeChars(string input)
        {
            return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
        }

    }

}
