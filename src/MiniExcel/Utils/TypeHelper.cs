namespace MiniExcelLibs.Utils
{
    using MiniExcelLibs.Exceptions;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Reflection;

    internal static partial class TypeHelper
    {
        public static IEnumerable<IDictionary<string, object>> ConvertToEnumerableDictionary(IDataReader reader)
        {
            while (reader.Read())
            {
                yield return Enumerable.Range(0, reader.FieldCount)
                 .ToDictionary(
                     i => reader.GetName(i),
                     i => reader.GetValue(i));
            }
        }


        /// <summary>
        /// From : https://stackoverflow.com/questions/906499/getting-type-t-from-ienumerablet
        /// </summary>
        public static IEnumerable<Type> GetGenericIEnumerables(object o)
        {
            return o.GetType()
                    .GetInterfaces()
                    .Where(t => t.IsGenericType
                        && t.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                    .Select(t => t.GetGenericArguments()[0]);
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

        public static object TypeMapping<T>(T v, ExcelColumnInfo pInfo, object newValue, object itemValue, int rowIndex, string startCell, Configuration _config) where T : class, new()
        {
            try
            {
                return TypeMappingImpl(v, pInfo, ref newValue, itemValue, _config);
            }
            catch (Exception ex) when (ex is InvalidCastException || ex is FormatException)
            {
                var columnName = pInfo.ExcelColumnName ?? pInfo.Property.Name;
                var startRowIndex = ReferenceHelper.ConvertCellToXY(startCell).Item2;
                var errorRow = startRowIndex + rowIndex + 1;
                throw new ExcelInvalidCastException(columnName, errorRow, itemValue, pInfo.Property.Info.PropertyType, $"ColumnName : {columnName}, CellRow : {errorRow}, Value : {itemValue}, it can't cast to {pInfo.Property.Info.PropertyType.Name} type.");
            }
        }

        private static object TypeMappingImpl<T>(T v, ExcelColumnInfo pInfo, ref object newValue, object itemValue, Configuration _config) where T : class, new()
        {
            if (pInfo.Nullable && string.IsNullOrWhiteSpace(itemValue?.ToString()))
            {
                newValue = null;
            }
            else if (pInfo.ExcludeNullableType == typeof(Guid))
            {
                newValue = Guid.Parse(itemValue.ToString());
            }
            else if (pInfo.ExcludeNullableType == typeof(DateTimeOffset))
            {
                var vs = itemValue?.ToString();
                if (pInfo.ExcelFormat != null)
                {
                    if (DateTimeOffset.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    {
                        newValue = _v;
                    }
                }
                else if (DateTimeOffset.TryParse(vs, _config.Culture, DateTimeStyles.None, out var _v))
                    newValue = _v;
                else
                    throw new InvalidCastException($"{vs} can't cast to datetime");
            }
            else if (pInfo.ExcludeNullableType == typeof(DateTime))
            {
                // fix issue 257 https://github.com/shps951023/MiniExcel/issues/257
                if (itemValue is DateTime || itemValue is DateTime?)
                {
                    newValue = itemValue;
                    pInfo.Property.SetValue(v, newValue);
                    return newValue;
                }

                var vs = itemValue?.ToString();
                if (pInfo.ExcelFormat != null)
                {
                    if (pInfo.Property.Info.PropertyType == typeof(DateTimeOffset) && DateTimeOffset.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                    {
                        newValue = _v2;
                    }
                    else if (DateTime.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    {
                        newValue = _v;
                    }
                }
                else if (DateTime.TryParse(vs, _config.Culture, DateTimeStyles.None, out var _v))
                    newValue = _v;
                else if (DateTime.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                    newValue = _v2;
                else if (double.TryParse(vs, NumberStyles.None, CultureInfo.InvariantCulture, out var _d))
                    newValue = DateTimeHelper.FromOADate(_d);
                else
                    throw new InvalidCastException($"{vs} can't cast to datetime");
            }
            else if (pInfo.ExcludeNullableType == typeof(bool))
            {
                var vs = itemValue.ToString();
                if (vs == "1")
                    newValue = true;
                else if (vs == "0")
                    newValue = false;
                else
                    newValue = bool.Parse(vs);
            }
            else if (pInfo.Property.Info.PropertyType == typeof(string))
            {
                newValue = XmlEncoder.DecodeString(itemValue?.ToString());
            }
            else if (pInfo.ExcludeNullableType.IsEnum)
            {
                var fieldInfo = pInfo.ExcludeNullableType.GetFields().FirstOrDefault(e => e.GetCustomAttribute<DescriptionAttribute>(false)?.Description == itemValue?.ToString());
                if (fieldInfo != null)
                    newValue = Enum.Parse(pInfo.ExcludeNullableType, fieldInfo.Name, true);
                else
                    newValue = Enum.Parse(pInfo.ExcludeNullableType, itemValue?.ToString(), true);
            }
            else
            {
                // Use pInfo.ExcludeNullableType to resolve : https://github.com/shps951023/MiniExcel/issues/138
                newValue = Convert.ChangeType(itemValue, pInfo.ExcludeNullableType, _config.Culture);
            }

            pInfo.Property.SetValue(v, newValue);
            return newValue;
        }

    }
}
