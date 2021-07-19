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

    internal static partial class TypeHelper
    {
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


        public static object TypeMapping<T>(T v, ExcelCustomPropertyInfo pInfo, object newValue, object itemValue, int rowIndex, string startCell) where T : class, new()
        {
            try
            {
                return TypeMappingImpl(v, pInfo, ref newValue, itemValue);
            }
            catch (Exception ex) when (ex is InvalidCastException || ex is FormatException)
            {
                var columnName = pInfo.ExcelColumnName ?? pInfo.Property.Name;
                var startRowIndex = ReferenceHelper.ConvertCellToXY(startCell).Item2;
                var errorRow = startRowIndex + rowIndex + 1;
                throw new InvalidCastException($"ColumnName : {columnName}, CellRow : {errorRow}, Value : {itemValue}, it can't cast to {pInfo.Property.PropertyType.Name} type.");
            }
        }

        private static object TypeMappingImpl<T>(T v, ExcelCustomPropertyInfo pInfo, ref object newValue, object itemValue) where T : class, new()
        {
            if (pInfo.ExcludeNullableType == typeof(Guid))
            {
                newValue = Guid.Parse(itemValue.ToString());
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
                    if (DateTime.TryParseExact(vs, pInfo.ExcelFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
                    {
                        newValue = _v;
                    }
                }
                else if (DateTime.TryParse(vs, CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v))
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
            else if (pInfo.Property.PropertyType == typeof(string))
            {
                newValue = XmlEncoder.DecodeString(itemValue?.ToString());
            }
            else if (pInfo.Property.PropertyType.IsEnum)
            {
                newValue = Enum.Parse(pInfo.Property.PropertyType, itemValue?.ToString(), true);
            }
            else
            {
                // Use pInfo.ExcludeNullableType to resolve : https://github.com/shps951023/MiniExcel/issues/138
                newValue = Convert.ChangeType(itemValue, pInfo.ExcludeNullableType);
            }

            pInfo.Property.SetValue(v, newValue);
            return newValue;
        }

    }
}
