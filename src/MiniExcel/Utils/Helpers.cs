namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.Attributes;
    using System;
    using System.Collections.Generic;
    using System.Dynamic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    internal static partial class Helpers
    {
        public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }

    // For Row/Column Index
    internal static partial class Helpers
    {
        private const int GENERAL_COLUMN_INDEX = 255;
        private const int MAX_COLUMN_INDEX = 16383;
        private static Dictionary<int, string> _IntMappingAlphabet;
        private static Dictionary<string, int> _AlphabetMappingInt;
        static Helpers()
        {
            if (_IntMappingAlphabet == null && _AlphabetMappingInt == null)
            {
                _IntMappingAlphabet = new Dictionary<int, string>();
                _AlphabetMappingInt = new Dictionary<string, int>();
                for (int i = 0; i <= GENERAL_COLUMN_INDEX; i++)
                {
                    _IntMappingAlphabet.Add(i, IntToLetters(i));
                    _AlphabetMappingInt.Add(IntToLetters(i), i);
                }
            }
        }

        public static string GetAlphabetColumnName(int columnIndex)
        {
            CheckAndSetMaxColumnIndex(columnIndex);
            return _IntMappingAlphabet[columnIndex];
        }

        public static int GetColumnIndex(string columnName)
        {
            var columnIndex = _AlphabetMappingInt[columnName];
            CheckAndSetMaxColumnIndex(columnIndex);
            return columnIndex;
        }

        private static void CheckAndSetMaxColumnIndex(int columnIndex)
        {
            if (columnIndex >= _IntMappingAlphabet.Count)
            {
                if (columnIndex > MAX_COLUMN_INDEX)
                    throw new InvalidDataException($"ColumnIndex {columnIndex} over excel vaild max index.");
                for (int i = _IntMappingAlphabet.Count; i <= columnIndex; i++)
                {
                    _IntMappingAlphabet.Add(i, IntToLetters(i));
                    _AlphabetMappingInt.Add(IntToLetters(i), i);
                }
            }
        }

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
    }

    internal static partial class Helpers
    {
        internal static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex, int startCellIndex)
        {
            // TODO: strong type mapping can ignore this
            // TODO: it can recode better performance 
            var cell = (IDictionary<string, object>)new ExpandoObject();
            for (int i = startCellIndex; i <= maxColumnIndex; i++)
            {
                var key = GetAlphabetColumnName(i);
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

        internal class ExcelCustomPropertyInfo
        {
            public int? ExcelColumnIndex { get; set; }
            public string ExcelColumnName { get; set; }
            public PropertyInfo Property { get; set; }
            public Type ExcludeNullableType { get; set; }
            public bool Nullable { get; internal set; }
            public string ExcelFormat { get; internal set; }
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

        public static bool IsNumericType(Type type, bool isNullableUnderlyingType = false)
        {
            if (isNullableUnderlyingType)
                type = Nullable.GetUnderlyingType(type) ?? type;
            switch (Type.GetTypeCode(type))
            {
                //case TypeCode.Byte:
                //case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

        public static object TypeMapping<T>(T v, ExcelCustomPropertyInfo pInfo, object newV, object itemValue) where T : class, new()
        {
            if (pInfo.ExcludeNullableType == typeof(Guid))
            {
                newV = Guid.Parse(itemValue.ToString());
            }
            else if (pInfo.ExcludeNullableType == typeof(DateTime))
            {
                // fix issue 257 https://github.com/shps951023/MiniExcel/issues/257
                if (itemValue is DateTime || itemValue is DateTime?)
                {
                    newV = itemValue;
                    pInfo.Property.SetValue(v, newV);
                    return newV;
                }

                var vs = itemValue?.ToString();
                if (pInfo.ExcelFormat != null)
                {
                    if (DateTime.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    {
                        newV = _v;
                    }
                }
                else if (DateTime.TryParse(vs, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    newV = _v;
                else if (DateTime.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                    newV = _v2;
                else if (double.TryParse(vs, NumberStyles.None, CultureInfo.InvariantCulture, out var _d))
                    newV = DateTimeHelper.FromOADate(_d);
                else
                    throw new InvalidCastException($"{vs} can't cast to datetime");
            }
            else if (pInfo.ExcludeNullableType == typeof(bool))
            {
                var vs = itemValue.ToString();
                if (vs == "1")
                    newV = true;
                else if (vs == "0")
                    newV = false;
                else
                    newV = bool.Parse(vs);
            }
            else if (pInfo.Property.PropertyType == typeof(string))
            {
                newV = XmlEncoder.DecodeString(itemValue?.ToString());
            }
            else if (pInfo.Property.PropertyType.IsEnum)
            {
                newV = Enum.Parse(pInfo.Property.PropertyType, itemValue?.ToString(), true);
            }
            // solve : https://github.com/shps951023/MiniExcel/issues/138
            else
                newV = Convert.ChangeType(itemValue, pInfo.ExcludeNullableType);
            pInfo.Property.SetValue(v, newV);
            return newV;
        }

    }

}
