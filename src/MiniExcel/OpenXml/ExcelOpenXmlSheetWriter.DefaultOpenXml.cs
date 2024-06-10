using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using static MiniExcelLibs.Utils.ImageHelper;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        public void Insert()
        {
            throw new NotImplementedException();
        }

        private readonly Dictionary<string, ZipPackageInfo> _zipDictionary = new Dictionary<string, ZipPackageInfo>();

        private IEnumerable<Tuple<SheetDto, object>> GetSheets()
        {
            var sheetId = 0;
            if (_value is IDictionary<string, object> dictionary)
            {
                foreach (var sheet in dictionary)
                {
                    sheetId++;
                    var sheetInfos = GetSheetInfos(sheet.Key);
                    yield return Tuple.Create(sheetInfos.ToDto(sheetId), sheet.Value);
                }

                yield break;
            }

            if (_value is DataSet dataSet)
            {
                foreach (DataTable dt in dataSet.Tables)
                {
                    sheetId++;
                    var sheetInfos = GetSheetInfos(dt.TableName);
                    yield return Tuple.Create<SheetDto, object>(sheetInfos.ToDto(sheetId), dt);
                }

                yield break;
            }

            sheetId++;
            var defaultSheetInfo = GetSheetInfos(_defaultSheetName);
            yield return Tuple.Create(defaultSheetInfo.ToDto(sheetId), _value);
        }

        private ExcellSheetInfo GetSheetInfos(string sheetName)
        {
            var info = new ExcellSheetInfo
            {
                ExcelSheetName = sheetName,
                Key = sheetName,
                ExcelSheetState = SheetState.Visible
            };

            if (_configuration.DynamicSheets == null || _configuration.DynamicSheets.Length <= 0)
            {
                return info;
            }

            var dynamicSheet = _configuration.DynamicSheets.SingleOrDefault(_ => _.Key == sheetName);
            if (dynamicSheet == null)
            {
                return info;
            }

            info.ExcelSheetState = dynamicSheet.State;
            if (dynamicSheet.Name != null)
            {
                info.ExcelSheetName = dynamicSheet.Name;
            }

            return info;
        }

        private ExcelColumnInfo GetColumnInfosFromDynamicConfiguration(string columnName)
        {
            var prop = new ExcelColumnInfo
            {
                ExcelColumnName = columnName,
                Key = columnName
            };

            if (_configuration.DynamicColumns == null || _configuration.DynamicColumns.Length <= 0)
                return prop;

            var dynamicColumn = _configuration.DynamicColumns.SingleOrDefault(_ => _.Key == columnName);
            if (dynamicColumn == null || dynamicColumn.Ignore)
            {
                return prop;
            }

            prop.Nullable = true;
            prop.ExcelIgnore = dynamicColumn.Ignore;
            prop.ExcelColumnIndex = dynamicColumn.Index;
            prop.ExcelColumnWidth = dynamicColumn.Width;
            //prop.ExcludeNullableType = item2[key]?.GetType();

            if (dynamicColumn.Format != null)
            {
                prop.ExcelFormat = dynamicColumn.Format;
                prop.ExcelFormatId = dynamicColumn.FormatId;
            }

            if (dynamicColumn.Aliases != null)
            {
                prop.ExcelColumnAliases = dynamicColumn.Aliases;
            }

            if (dynamicColumn.IndexName != null)
            {
                prop.ExcelIndexName = dynamicColumn.IndexName;
            }

            if (dynamicColumn.Name != null)
            {
                prop.ExcelColumnName = dynamicColumn.Name;
            }

            return prop;
        }

        private void SetGenericTypePropertiesMode(Type genericType, ref string mode, out int maxColumnIndex, out List<ExcelColumnInfo> props)
        {
            mode = "Properties";
            if (genericType.IsValueType)
            {
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} value generic type");
            }
            
            if (genericType == typeof(string) || genericType == typeof(DateTime) || genericType == typeof(Guid))
            {
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} generic type");
            }

            props = CustomPropertyHelper.GetSaveAsProperties(genericType, _configuration);

            maxColumnIndex = props.Count;
        }

        private Tuple<string, string, string> GetCellValue(int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo, bool valueIsNull)
        {
            if (valueIsNull)
            {
                return Tuple.Create("2", "str", string.Empty);
            }

            if (value is string str)
            {
                return Tuple.Create("2", "str", ExcelOpenXmlUtils.EncodeXML(str));
            }

            var type = GetValueType(value, columnInfo);

            if (type != typeof(DateTime) && columnInfo?.ExcelFormat != null && value is IFormattable formattableValue)
            {
                var formattedStr = formattableValue.ToString(columnInfo.ExcelFormat, _configuration.Culture);
                return Tuple.Create("2", "str", ExcelOpenXmlUtils.EncodeXML(formattedStr));
            }

            if (type == typeof(DateTime))
            {
                return GetDateTimeValue(value, columnInfo);
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly))
            {
                if (_configuration.Culture != CultureInfo.InvariantCulture)
                {
                    var cellValue = ((DateOnly)value).ToString(_configuration.Culture);
                    return Tuple.Create("2", "str", cellValue);
                }

                if (columnInfo == null || columnInfo.ExcelFormat == null)
                {
                    var oaDate = CorrectDateTimeValue((DateTime)value);
                    var cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
                    return Tuple.Create<string, string, string>("3", null, cellValue);
                }

                // TODO: now it'll lose date type information
                var formattedCellValue = ((DateOnly)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                return Tuple.Create("2", "str", formattedCellValue);
            }
#endif
            if (type.IsEnum)
            {
                var description = CustomPropertyHelper.DescriptionAttr(type, value);
                return Tuple.Create("2", "str", description ?? value.ToString());
            }
            
            if (TypeHelper.IsNumericType(type))
            {
                var dataType = _configuration.Culture == CultureInfo.InvariantCulture ? "n" : "str";
                string cellValue;

                cellValue = GetNumericValue(value, type);

                return Tuple.Create("2", dataType, cellValue);
            }
            
            if (type == typeof(bool))
            {
                return Tuple.Create("2", "b", (bool)value ? "1" : "0");
            }

            if (type == typeof(byte[]) && _configuration.EnableConvertByteArray)
            {
                var base64 = GetFileValue(rowIndex, cellIndex, value);
                return Tuple.Create("4", "str", ExcelOpenXmlUtils.EncodeXML(base64));
            }

            if (type == typeof(DateTime))
            {
                return GetDateTimeValue(value, columnInfo);
            }

#if NET6_0_OR_GREATER
            if (type == typeof(DateOnly))
            {
                if (_configuration.Culture != CultureInfo.InvariantCulture)
                {
                    var cellValue = ((DateOnly)value).ToString(_configuration.Culture);
                    return Tuple.Create("2", "str", cellValue);
                }

                if (columnInfo == null || columnInfo.ExcelFormat == null)
                {
                    var oaDate = CorrectDateTimeValue((DateTime)value);
                    var cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
                    return Tuple.Create<string, string, string>("3", null, cellValue);
                }

                // TODO: now it'll lose date type information
                var formattedCellValue = ((DateOnly)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                return Tuple.Create("2", "str", formattedCellValue);
            }
#endif

            return Tuple.Create("2", "str", ExcelOpenXmlUtils.EncodeXML(value.ToString()));
        }

        private static Type GetValueType(object value, ExcelColumnInfo columnInfo)
        {
            Type type;
            if (columnInfo == null || columnInfo.Key != null)
            {
                // TODO: need to optimize
                // Dictionary need to check type every time, so it's slow..
                type = value.GetType();
                type = Nullable.GetUnderlyingType(type) ?? type;
            }
            else
            {
                type = columnInfo.ExcludeNullableType; //sometime it doesn't need to re-get type like prop
            }

            return type;
        }

        private string GetNumericValue(object value, Type type)
        {
            if (type.IsAssignableFrom(typeof(decimal)))
            {
                return ((decimal)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(int)))
            {
                return ((int)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(double)))
            {
                return ((double)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(long)))
            {
                return ((long)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(uint)))
            {
                return ((uint)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(ushort)))
            {
                return ((ushort)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(ulong)))
            {
                return ((ulong)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(short)))
            {
                return ((short)value).ToString(_configuration.Culture);
            }
            
            if (type.IsAssignableFrom(typeof(float)))
            {
                return ((float)value).ToString(_configuration.Culture);
            }

            return (decimal.Parse(value.ToString())).ToString(_configuration.Culture);
        }

        private string GetFileValue(int rowIndex, int cellIndex, object value)
        {
            var bytes = (byte[])value;

            // TODO: Setting configuration because it might have high cost?
            var format = GetImageFormat(bytes);
            //it can't insert to zip first to avoid cache image to memory
            //because sheet xml is opening.. https://github.com/shps951023/MiniExcel/issues/304#issuecomment-1017031691
            //int rowIndex, int cellIndex
            var file = new FileDto()
            {
                Byte = bytes,
                RowIndex = rowIndex,
                CellIndex = cellIndex,
                SheetId = currentSheetIndex
            };

            if (format != ImageFormat.unknown)
            {
                file.Extension = format.ToString();
                file.IsImage = true;
            }
            else
            {
                file.Extension = "bin";
            }

            _files.Add(file);

            //TODO:Convert to base64
            var base64 = $"@@@fileid@@@,{file.Path}";
            return base64;
        }

        private Tuple<string, string, string> GetDateTimeValue(object value, ExcelColumnInfo columnInfo)
        {
            if (_configuration.Culture != CultureInfo.InvariantCulture)
            {
                var cellValue = ((DateTime)value).ToString(_configuration.Culture);
                return Tuple.Create("2", "str", cellValue);
            }

            if (columnInfo == null || columnInfo.ExcelFormat == null)
            {
                var oaDate = CorrectDateTimeValue((DateTime)value);
                var cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
                return Tuple.Create<string, string, string>("3", null, cellValue);
            }

            return Tuple.Create(columnInfo.ExcelFormatId.ToString(), (string)null, ((DateTime)value).ToOADate().ToString());
        }

        private static double CorrectDateTimeValue(DateTime value)
        {
            var oaDate = value.ToOADate();

            // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
            // to Lotus 1-2-3 decision from 1983...
            // https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Extensions/DateTimeExtensions.cs#L45
            const int nonExistent1900Feb29SerialDate = 60;
            if (oaDate <= nonExistent1900Feb29SerialDate)
            {
                oaDate -= 1;
            }

            return oaDate;
        }

        private string GetDimensionRef(int maxRowIndex, int maxColumnIndex)
        {
            string dimensionRef;
            if (maxRowIndex == 0 && maxColumnIndex == 0)
                dimensionRef = "A1";
            else if (maxColumnIndex == 1)
                dimensionRef = $"A{maxRowIndex}";
            else if (maxRowIndex == 0)
                dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}1";
            else
                dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}{maxRowIndex}";
            return dimensionRef;
        }

        private string GetStylesXml(ICollection<ExcelColumnAttribute> columns)
        {
            switch (_configuration.TableStyles)
            {
                case TableStyles.None:
                    return ExcelXml.SetupStyleXml(ExcelXml.NoneStylesXml, columns);
                case TableStyles.Default:
                    return ExcelXml.SetupStyleXml(ExcelXml.DefaultStylesXml, columns);
                default:
                    return string.Empty;
            }
        }

        private string GetDrawingRelationshipXml(int sheetIndex)
        {
            var drawing = new StringBuilder();
            foreach (var image in _files.Where(w => w.IsImage && w.SheetId == sheetIndex + 1))
            {
                drawing.AppendLine(ExcelXml.ImageRelationship(image));
            }

            return drawing.ToString();
        }

        private string GetDrawingXml(int sheetIndex)
        {
            var drawing = new StringBuilder();

            for (int fileIndex = 0; fileIndex < _files.Count; fileIndex++)
            { 
                var file = _files[fileIndex];
                if (file.IsImage && file.SheetId == sheetIndex + 1)
                {
                    drawing.Append(ExcelXml.DrawingXml(file, fileIndex));
                }
            }

            return drawing.ToString();
        }

        private void GenerateWorkBookXmls(
            out StringBuilder workbookXml,
            out StringBuilder workbookRelsXml,
            out Dictionary<int, string> sheetsRelsXml)
        {
            workbookXml = new StringBuilder();
            workbookRelsXml = new StringBuilder();
            sheetsRelsXml = new Dictionary<int, string>();
            var sheetId = 0;
            foreach (var sheetDto in _sheets)
            {
                sheetId++;
                workbookXml.AppendLine(ExcelXml.Sheet(sheetDto, sheetId));

                workbookRelsXml.AppendLine(ExcelXml.WorksheetRelationship(sheetDto));

                //TODO: support multiple drawing
                //TODO: ../drawings/drawing1.xml or /xl/drawings/drawing1.xml
                sheetsRelsXml.Add(sheetDto.SheetIdx, ExcelXml.DrawingRelationship(sheetId));
            }
        }

        private string GetContentTypesXml()
        {
            var sb = new StringBuilder(ExcelXml.StartTypes);
            foreach (var p in _zipDictionary)
            {
                sb.Append(ExcelXml.ContentType(p.Value.ContentType, p.Key));
            }

            sb.Append(ExcelXml.EndTypes);
            return sb.ToString();
        }
    }
}
