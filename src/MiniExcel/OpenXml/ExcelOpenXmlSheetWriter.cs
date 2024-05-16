using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using static MiniExcelLibs.Utils.ImageHelper;

namespace MiniExcelLibs.OpenXml
{
    internal class FileDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Extension { get; set; }
        public string Path { get { return $"xl/media/{ID}.{Extension}"; } }
        public string Path2 { get { return $"/xl/media/{ID}.{Extension}"; } }
        public Byte[] Byte { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
        public bool IsImage { get; set; } = false;
        public int SheetId { get; set; }
    }
    internal class SheetDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
        public string Name { get; set; }
        public int SheetIdx { get; set; }
        public string Path { get { return $"xl/worksheets/sheet{SheetIdx}.xml"; } }

        public string State { get; set; }
    }
    internal class DrawingDto
    {
        public string ID { get; set; } = $"R{Guid.NewGuid():N}";
    }
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly MiniExcelZipArchive _archive;
        private readonly static UTF8Encoding _utf8WithBom = new System.Text.UTF8Encoding(true);
        private readonly OpenXmlConfiguration _configuration;
        private readonly Stream _stream;
        private readonly bool _printHeader;
        private readonly object _value;
        private readonly List<SheetDto> _sheets = new List<SheetDto>();
        private readonly List<FileDto> _files = new List<FileDto>();
        private int currentSheetIndex = 0;

        public ExcelOpenXmlSheetWriter(Stream stream, object value, string sheetName, IConfiguration configuration, bool printHeader)
        {
            this._stream = stream;
            // Why ZipArchiveMode.Update not ZipArchiveMode.Create?
            // R : Mode create - ZipArchiveEntry does not support seeking.'
            this._configuration = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.DefaultConfig;
            if (_configuration.FastMode)
                this._archive = new MiniExcelZipArchive(_stream, ZipArchiveMode.Update, true, _utf8WithBom);
            else
                this._archive = new MiniExcelZipArchive(_stream, ZipArchiveMode.Create, true, _utf8WithBom);
            this._printHeader = printHeader;
            this._value = value;
            var defaultSheetInfo = GetSheetInfos(sheetName);
            _sheets.Add(defaultSheetInfo.ToDto(1)); //TODO:remove
        }

        public ExcelOpenXmlSheetWriter()
        {
        }

        public void SaveAs()
        {
            GenerateDefaultOpenXml();
            {
                if (_value is IDictionary<string, object>)
                {
                    var sheetId = 0;
                    var sheets = _value as IDictionary<string, object>;
                    _sheets.RemoveAt(0);//TODO:remove
                    foreach (var sheet in sheets)
                    {
                        sheetId++;
                        var sheetInfos = GetSheetInfos(sheet.Key);
                        var sheetDto = sheetInfos.ToDto(sheetId);
                        _sheets.Add(sheetDto); //TODO:remove

                        currentSheetIndex = sheetId;

                        CreateSheetXml(sheet.Value, sheetDto.Path);
                    }
                }
                else if (_value is DataSet)
                {
                    var sheetId = 0;
                    var sheets = _value as DataSet;
                    _sheets.RemoveAt(0);//TODO:remove
                    foreach (DataTable dt in sheets.Tables)
                    {
                        sheetId++;
                        var sheetInfos = GetSheetInfos(dt.TableName);
                        var sheetDto = sheetInfos.ToDto(sheetId);
                        _sheets.Add(sheetDto); //TODO:remove

                        currentSheetIndex = sheetId;

                        CreateSheetXml(dt, sheetDto.Path);
                    }
                }
                else
                {
                    //Single sheet export.
                    currentSheetIndex++;

                    CreateSheetXml(_value, _sheets[0].Path);
                }
            }
            GenerateEndXml();
            _archive.Dispose();
        }

        private void GenerateSheetByEnumerable(MiniExcelStreamWriter writer, IEnumerable values)
        {
            var maxColumnIndex = 0;
            var maxRowIndex = 0;
            List<ExcelColumnInfo> props = null;
            string mode = null;

            int? rowCount = null;
            var collection = values as ICollection;
            if (collection != null)
            {
                rowCount = collection.Count;
            }
            else if (!_configuration.FastMode)
            {
                // The row count is only required up front when not in fastmode
                collection = new List<object>(values.Cast<object>());
                rowCount = collection.Count;
            }

            // Get the enumerator once to ensure deferred linq execution
            var enumerator = (collection ?? values).GetEnumerator();

            // Move to the first item in order to inspect the value type and determine whether it is empty
            var empty = !enumerator.MoveNext();

            if (empty)
            {
                // only when empty IEnumerable need to check this issue #133  https://github.com/shps951023/MiniExcel/issues/133
                var genericType = TypeHelper.GetGenericIEnumerables(values).FirstOrDefault();
                if (genericType == null || genericType == typeof(object) // sometime generic type will be object, e.g: https://user-images.githubusercontent.com/12729184/132812859-52984314-44d1-4ee8-9487-2d1da159f1f0.png
                    || typeof(IDictionary<string, object>).IsAssignableFrom(genericType)
                    || typeof(IDictionary).IsAssignableFrom(genericType))
                {
                    WriteEmptySheet(writer);
                    return;
                }
                else
                {
                    SetGenericTypePropertiesMode(genericType, ref mode, out maxColumnIndex, out props);
                }
            }
            else
            {
                var firstItem = enumerator.Current;
                if (firstItem is IDictionary<string, object> genericDic)
                {
                    mode = "IDictionary<string, object>";
                    props = GetDictionaryColumnInfo(genericDic, null);
                    maxColumnIndex = props.Count;
                }
                else if (firstItem is IDictionary dic)
                {
                    mode = "IDictionary";
                    props = GetDictionaryColumnInfo(null, dic);
                    //maxColumnIndex = dic.Keys.Count;
                    maxColumnIndex = props.Count; // why not using keys, because ignore attribute ![image](https://user-images.githubusercontent.com/12729184/163686902-286abb70-877b-4e84-bd3b-001ad339a84a.png)
                }
                else
                {
                    SetGenericTypePropertiesMode(firstItem.GetType(), ref mode, out maxColumnIndex, out props);
                }
            }

            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" >");

            long dimensionWritePosition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && rowCount == null)
            {
                // Write a placeholder for the table dimensions and save thee position for later
                dimensionWritePosition = writer.WriteAndFlush("<x:dimension ref=\"");
                writer.Write("                              />");
            }
            else
            {
                maxRowIndex = rowCount.Value + (_printHeader && rowCount > 0 ? 1 : 0);
                writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/>");
            }

            //cols:width
            WriteColumnsWidths(writer, props);

            //header
            writer.Write($@"<x:sheetData>");
            var yIndex = 1;
            var xIndex = 1;
            if (_printHeader)
            {
                PrintHeader(writer, props);
                yIndex++;
            }

            if (!empty)
            {
                // body
                if (mode == "IDictionary<string, object>") //Dapper Row
                    maxRowIndex = GenerateSheetByColumnInfo<IDictionary<string, object>>(writer, enumerator, props, xIndex, yIndex);
                else if (mode == "IDictionary") //IDictionary
                    maxRowIndex = GenerateSheetByColumnInfo<IDictionary>(writer, enumerator, props, xIndex, yIndex);
                else if (mode == "Properties")
                    maxRowIndex = GenerateSheetByColumnInfo<object>(writer, enumerator, props, xIndex, yIndex);
                else
                    throw new NotImplementedException($"Type {values.GetType().FullName} is not implemented. Please open an issue.");
            }

            writer.Write("</x:sheetData>");
            if (_configuration.AutoFilter)
                writer.Write($"<x:autoFilter ref=\"{GetDimensionRef(maxRowIndex, maxColumnIndex)}\" />");

            // The dimension has already been written if row count is defined
            if (_configuration.FastMode && rowCount == null)
            {
                // Flush and save position so that we can get back again.
                var pos = writer.Flush();

                // Seek back and write the dimensions of the table
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
                writer.SetPosition(pos);
            }

            writer.Write("<x:drawing  r:id=\"drawing" + currentSheetIndex + "\" /></x:worksheet>");
        }

        private static void PrintHeader(MiniExcelStreamWriter writer, List<ExcelColumnInfo> props)
        {
            var xIndex = 1;
            var yIndex = 1;
            writer.Write($"<x:row r=\"{yIndex}\">");

            foreach (var p in props)
            {
                if (p == null)
                {
                    xIndex++; //reason : https://github.com/shps951023/MiniExcel/issues/142
                    continue;
                }

                var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                WriteC(writer, r, columnName: p.ExcelColumnName);
                xIndex++;
            }

            writer.Write("</x:row>");
        }

        private void CreateSheetXml(object value, string sheetPath)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
            {
                if (value == null)
                {
                    WriteEmptySheet(writer);
                    goto End; //for re-using code
                }

                //DapperRow

                if (value is IDataReader)
                {
                    GenerateSheetByIDataReader(writer, value as IDataReader);
                }
                else if (value is IEnumerable)
                {
                    GenerateSheetByEnumerable(writer, value as IEnumerable);
                }
                else if (value is DataTable)
                {
                    GenerateSheetByDataTable(writer, value as DataTable);
                }
                else
                {
                    throw new NotImplementedException($"Type {value.GetType().FullName} is not implemented. Please open an issue.");
                }
            }
        End: //for re-using code
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
        }

        private List<ExcelColumnInfo> GetDictionaryColumnInfo(IDictionary<string, object> dicString, IDictionary dic)
        {
            List<ExcelColumnInfo> props;
            var _props = new List<ExcelColumnInfo>();
            if (dicString != null)
                foreach (var key in dicString.Keys)
                    SetDictionaryColumnInfo(_props, key);
            else if (dic != null)
                foreach (var key in dic.Keys)
                    SetDictionaryColumnInfo(_props, key);
            else
                throw new NotSupportedException("SetDictionaryColumnInfo Error");
            props = CustomPropertyHelper.SortCustomProps(_props);
            return props;
        }

        private void SetDictionaryColumnInfo(List<ExcelColumnInfo> _props, object key)
        {
            var p = new ExcelColumnInfo();
            p.ExcelColumnName = key?.ToString();
            p.Key = key;
            // TODO:Dictionary value type is not fiexed
            //var _t =
            //var gt = Nullable.GetUnderlyingType(p.PropertyType);
            var isIgnore = false;
            if (_configuration.DynamicColumns != null && _configuration.DynamicColumns.Length > 0)
            {
                var dynamicColumn = _configuration.DynamicColumns.SingleOrDefault(_ => _.Key == key.ToString());
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

        private void SetGenericTypePropertiesMode(Type genericType, ref string mode, out int maxColumnIndex, out List<ExcelColumnInfo> props)
        {
            mode = "Properties";
            if (genericType.IsValueType)
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} value generic type");
            else if (genericType == typeof(string) || genericType == typeof(DateTime) || genericType == typeof(Guid))
                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} generic type");
            props = CustomPropertyHelper.GetSaveAsProperties(genericType, _configuration);

            maxColumnIndex = props.Count;
        }

        private void WriteEmptySheet(MiniExcelStreamWriter writer)
        {
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><x:dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");
        }

        private int GenerateSheetByColumnInfo<T>(MiniExcelStreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, int xIndex = 1, int yIndex = 1)
        {
            var isDic = typeof(T) == typeof(IDictionary);
            var isDapperRow = typeof(T) == typeof(IDictionary<string, object>);
            do
            {
                // The enumerator has already moved to the first item
                T v = (T)value.Current;

                writer.Write($"<x:row r=\"{yIndex}\">");
                var cellIndex = xIndex;
                foreach (var columnInfo in props)
                {
                    if (columnInfo == null) //reason:https://github.com/shps951023/MiniExcel/issues/142
                    {
                        cellIndex++;
                        continue;
                    }
                    object cellValue = null;
                    if (isDic)
                    {
                        cellValue = ((IDictionary)v)[columnInfo.Key];
                        //WriteCell(writer, yIndex, cellIndex, cellValue, null); // why null because dictionary that needs to check type every time
                        //TODO: user can specefic type to optimize efficiency
                    }
                    else if (isDapperRow)
                    {
                        cellValue = ((IDictionary<string, object>)v)[columnInfo.Key.ToString()];
                    }
                    else
                    {
                        cellValue = columnInfo.Property.GetValue(v);
                    }

                    WriteCell(writer, yIndex, cellIndex, cellValue, columnInfo);

                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            } while (value.MoveNext());

            return yIndex - 1;
        }

        private void WriteCell(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo)
        {
            var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                writer.Write($"<x:c r=\"{columname}\" s=\"2\"></x:c>"); // s: style index
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

            var styleIndex = tuple.Item1; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
            var dataType = tuple.Item2; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1
            var cellValue = tuple.Item3;

            if (cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) || cellValue.EndsWith(" ", StringComparison.Ordinal))) /*Prefix and suffix blank space will lost after SaveAs #294*/
            {
                writer.Write($"<x:c r=\"{columname}\" {(dataType == null ? "" : $"t =\"{dataType}\"")} s=\"{styleIndex}\" xml:space=\"preserve\"><x:v>{cellValue}</x:v></x:c>");
            }
            else
            {
                //t check avoid format error ![image](https://user-images.githubusercontent.com/12729184/118770190-9eee3480-b8b3-11eb-9f5a-87a439f5e320.png)
                writer.Write($"<x:c r=\"{columname}\" {(dataType == null ? "" : $"t =\"{dataType}\"")} s=\"{styleIndex}\"><x:v>{cellValue}</x:v></x:c>");
            }
        }

        private Tuple<string, string, string> GetCellValue(int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo, bool valueIsNull)
        {
            var styleIndex = "2"; // format code: 0.00
            var cellValue = string.Empty;
            var dataType = "str";

            if (valueIsNull)
            {
                // use defaults
            }
            else if (value is string str)
            {
                cellValue = ExcelOpenXmlUtils.EncodeXML(str);
            }
            else if (columnInfo?.ExcelFormat != null && value is IFormattable formattableValue)
            {
                var formattedStr = formattableValue.ToString(columnInfo.ExcelFormat, _configuration.Culture);
                cellValue = ExcelOpenXmlUtils.EncodeXML(formattedStr);
            }
            else
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

                if (type.IsEnum)
                {
                    dataType = "str";
                    var description = CustomPropertyHelper.DescriptionAttr(type, value); //TODO: need to optimze
                    if (description != null)
                        cellValue = description;
                    else
                        cellValue = value.ToString();
                }
                else if (TypeHelper.IsNumericType(type))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                        dataType = "str"; //TODO: add style format
                    else
                        dataType = "n";

                    if (type.IsAssignableFrom(typeof(decimal)))
                        cellValue = ((decimal)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int32)))
                        cellValue = ((Int32)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Double)))
                        cellValue = ((Double)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int64)))
                        cellValue = ((Int64)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt32)))
                        cellValue = ((UInt32)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt16)))
                        cellValue = ((UInt16)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(UInt64)))
                        cellValue = ((UInt64)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Int16)))
                        cellValue = ((Int16)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        cellValue = ((Single)value).ToString(_configuration.Culture);
                    else if (type.IsAssignableFrom(typeof(Single)))
                        cellValue = ((Single)value).ToString(_configuration.Culture);
                    else
                        cellValue = (decimal.Parse(value.ToString())).ToString(_configuration.Culture);
                }
                else if (type == typeof(bool))
                {
                    dataType = "b";
                    cellValue = (bool)value ? "1" : "0";
                }
                else if (type == typeof(byte[]) && _configuration.EnableConvertByteArray)
                {
                    var bytes = (byte[])value;
                    if (bytes != null)
                    {
                        // TODO: Setting configuration because it might have high cost?
                        var format = ImageHelper.GetImageFormat(bytes);
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
                        cellValue = ExcelOpenXmlUtils.EncodeXML(base64);
                        styleIndex = "4";
                    }
                }
                else if (type == typeof(DateTime))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                    {
                        dataType = "str";
                        cellValue = ((DateTime)value).ToString(_configuration.Culture);
                    }
                    else if (columnInfo == null || columnInfo.ExcelFormat == null)
                    {
                        dataType = null;
                        styleIndex = "3";
                        cellValue = ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // TODO: now it'll lose date type information
                        dataType = "str";
                        cellValue = ((DateTime)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                    }
                }
#if NET6_0_OR_GREATER
                else if (type == typeof(DateOnly))
                {
                    if (_configuration.Culture != CultureInfo.InvariantCulture)
                    {
                        dataType = "str";
                        cellValue = ((DateOnly)value).ToString(_configuration.Culture);
                    }
                    else if (columnInfo == null || columnInfo.ExcelFormat == null)
                    {
                        var day = (DateOnly)value;
                        dataType = "n";
                        styleIndex = "3";
                        cellValue = day.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        // TODO: now it'll lose date type information
                        dataType = "str";
                        cellValue = ((DateOnly)value).ToString(columnInfo.ExcelFormat, _configuration.Culture);
                    }
                }
#endif
                else
                {
                    //TODO: _configuration.Culture
                    cellValue = ExcelOpenXmlUtils.EncodeXML(value.ToString());
                }
            }

            return Tuple.Create(styleIndex, dataType, cellValue);
        }

        private void GenerateSheetByDataTable(MiniExcelStreamWriter writer, DataTable value)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                // dimension
                var maxRowIndex = value.Rows.Count + (_printHeader && value.Rows.Count > 0 ? 1 : 0);
                var maxColumnIndex = value.Columns.Count;
                writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/>");

                var props = new List<ExcelColumnInfo>();
                for (var i = 0; i < value.Columns.Count; i++)
                {
                    var columnName = value.Columns[i].Caption ?? value.Columns[i].ColumnName;
                    var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                    props.Add(prop);
                }

                WriteColumnsWidths(writer, props);

                writer.Write("<x:sheetData>");
                if (_printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        WriteC(writer, r, columnName: p.ExcelColumnName);
                        xIndex++;
                    }

                    writer.Write($"</x:row>");
                    yIndex++;
                }

                for (int i = 0; i < value.Rows.Count; i++)
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    var xIndex = xy.Item1;

                    for (int j = 0; j < value.Columns.Count; j++)
                    {
                        var cellValue = value.Rows[i][j];
                        WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: null);
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
            }
            writer.Write("</x:sheetData></x:worksheet>");
        }

        private void GenerateSheetByIDataReader(MiniExcelStreamWriter writer, IDataReader reader)
        {
            long dimensionWritePosition = 0;
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            var xIndex = 1;
            var yIndex = 1;
            var maxColumnIndex = 0;
            var maxRowIndex = 0;
            {

                if (_configuration.FastMode)
                {
                    dimensionWritePosition = writer.WriteAndFlush($@"<x:dimension ref=""");
                    writer.Write("                              />"); // end of code will be replaced
                }

                var props = new List<ExcelColumnInfo>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                    props.Add(prop);
                }
                maxColumnIndex = props.Count;

                WriteColumnsWidths(writer, props);

                writer.Write("<x:sheetData>");
                int fieldCount = reader.FieldCount;
                if (_printHeader)
                {
                    PrintHeader(writer, props);
                    yIndex++;
                }

                while (reader.Read())
                {
                    writer.Write($"<x:row r=\"{yIndex}\">");
                    xIndex = 1;
                    for (int i = 0; i < fieldCount; i++)
                    {
                        var cellValue = reader.GetValue(i);
                        WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: null);
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }

                // Subtract 1 because cell indexing starts with 1
                maxRowIndex = yIndex - 1;
            }
            writer.Write("</x:sheetData>");
            if (_configuration.AutoFilter)
                writer.Write($"<x:autoFilter ref=\"{GetDimensionRef(maxRowIndex, maxColumnIndex)}\" />");
            writer.WriteAndFlush("</x:worksheet>");

            if (_configuration.FastMode)
            {
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
            }
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
            //prop.ExcludeNullableType = item2[key]?.GetType();
            if (dynamicColumn.Format != null)
                prop.ExcelFormat = dynamicColumn.Format;
            if (dynamicColumn.Aliases != null)
                prop.ExcelColumnAliases = dynamicColumn.Aliases;
            if (dynamicColumn.IndexName != null)
                prop.ExcelIndexName = dynamicColumn.IndexName;
            prop.ExcelColumnIndex = dynamicColumn.Index;
            if (dynamicColumn.Name != null)
                prop.ExcelColumnName = dynamicColumn.Name;
            prop.ExcelColumnWidth = dynamicColumn.Width;

            return prop;
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
                return info;

            var dynamicSheet = _configuration.DynamicSheets.SingleOrDefault(_ => _.Key == sheetName);
            if (dynamicSheet == null)
            {
                return info;
            }

            if (dynamicSheet.Name != null)
                info.ExcelSheetName = dynamicSheet.Name;
            info.ExcelSheetState = dynamicSheet.State;

            return info;
        }

        private static void WriteColumnsWidths(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnInfo> props)
        {
            var ecwProps = props.Where(x => x?.ExcelColumnWidth != null).ToList();
            if (ecwProps.Count <= 0)
                return;
            writer.Write($@"<x:cols>");
            foreach (var p in ecwProps)
            {
                writer.Write(
                    $@"<x:col min=""{p.ExcelColumnIndex + 1}"" max=""{p.ExcelColumnIndex + 1}"" width=""{p.ExcelColumnWidth?.ToString(CultureInfo.InvariantCulture)}"" customWidth=""1"" />");
            }

            writer.Write($@"</x:cols>");
        }

        private static void WriteC(MiniExcelStreamWriter writer, string r, string columnName)
        {
            writer.Write($"<x:c r=\"{r}\" t=\"str\" s=\"1\">");
            writer.Write($"<x:v>{ExcelOpenXmlUtils.EncodeXML(columnName)}"); //issue I45TF5
            writer.Write($"</x:v>");
            writer.Write($"</x:c>");
        }

        private void GenerateEndXml()
        {
            //Files
            {
                foreach (var item in _files)
                {
                    this.CreateZipEntry(item.Path, item.Byte);
                }
            }

            // styles.xml
            {
                var styleXml = string.Empty;

                if (_configuration.TableStyles == TableStyles.None)
                {
                    styleXml = _noneStylesXml;
                }
                else if (_configuration.TableStyles == TableStyles.Default)
                {
                    styleXml = _defaultStylesXml;
                }
                CreateZipEntry(@"xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", styleXml);
            }

            // drawing rel
            {
                for (int j = 0; j < _sheets.Count; j++)
                {
                    var drawing = new StringBuilder();
                    foreach (var i in _files.Where(w => w.IsImage && w.SheetId == j + 1))
                    {
                        drawing.AppendLine($@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""{i.Path2}"" Id=""{i.ID}"" />");
                    }
                    CreateZipEntry($"xl/drawings/_rels/drawing{j + 1}.xml.rels", "",
                        _defaultDrawingXmlRels.Replace("{{format}}", drawing.ToString()));
                }

            }
            // drawing
            {
                for (int j = 0; j < _sheets.Count; j++)
                {
                    var drawing = new StringBuilder();
                    foreach (var i in _files.Where(w => w.IsImage && w.SheetId == j + 1))
                    {
                        drawing.Append($@"<xdr:oneCellAnchor>
        <xdr:from>
            <xdr:col>{i.CellIndex - 1/* why -1 : https://user-images.githubusercontent.com/12729184/150460189-f08ed939-44d4-44e1-be6e-9c533ece6be8.png*/}</xdr:col>
            <xdr:colOff>0</xdr:colOff>
            <xdr:row>{i.RowIndex - 1}</xdr:row>
            <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:ext cx=""609600"" cy=""190500"" />
        <xdr:pic>
            <xdr:nvPicPr>
                <xdr:cNvPr id=""{_files.IndexOf(i) + 1}"" descr="""" name=""2a3f9147-58ea-4a79-87da-7d6114c4877b"" />
                <xdr:cNvPicPr>
                    <a:picLocks noChangeAspect=""1"" />
                </xdr:cNvPicPr>
            </xdr:nvPicPr>
            <xdr:blipFill>
                <a:blip r:embed=""{i.ID}"" cstate=""print"" />
                <a:stretch>
                    <a:fillRect />
                </a:stretch>
            </xdr:blipFill>
            <xdr:spPr>
                <a:xfrm>
                    <a:off x=""0"" y=""0"" />
                    <a:ext cx=""0"" cy=""0"" />
                </a:xfrm>
                <a:prstGeom prst=""rect"">
                    <a:avLst />
                </a:prstGeom>
            </xdr:spPr>
        </xdr:pic>
        <xdr:clientData />
    </xdr:oneCellAnchor>");
                    }
                    CreateZipEntry($"xl/drawings/drawing{j + 1}.xml", "application/vnd.openxmlformats-officedocument.drawing+xml",
                        _defaultDrawing.Replace("{{format}}", drawing.ToString()));
                }
            }

            // workbook.xml 、 workbookRelsXml
            {
                var workbookXml = new StringBuilder();
                var workbookRelsXml = new StringBuilder();

                var sheetId = 0;
                foreach (var s in _sheets)
                {
                    sheetId++;
                    if (string.IsNullOrEmpty(s.State))
                    {
                        workbookXml.AppendLine($@"<x:sheet name=""{ExcelOpenXmlUtils.EncodeXML(s.Name)}"" sheetId=""{sheetId}"" r:id=""{s.ID}"" />");
                    }
                    else
                    {
                        workbookXml.AppendLine($@"<x:sheet name=""{ExcelOpenXmlUtils.EncodeXML(s.Name)}"" sheetId=""{sheetId}"" state=""{s.State}"" r:id=""{s.ID}"" />");
                    }
                    workbookRelsXml.AppendLine($@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/{s.Path}"" Id=""{s.ID}"" />");

                    //TODO: support multiple drawing
                    //TODO: ../drawings/drawing1.xml or /xl/drawings/drawing1.xml
                    var sheetRelsXml = $@"<Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"" Target=""../drawings/drawing{sheetId}.xml"" Id=""drawing{sheetId}"" />";
                    CreateZipEntry($"xl/worksheets/_rels/sheet{s.SheetIdx}.xml.rels", "",
                        _defaultSheetRelXml.Replace("{{format}}", sheetRelsXml));
                }
                CreateZipEntry(@"xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    _defaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()));
                CreateZipEntry(@"xl/_rels/workbook.xml.rels", "",
                    _defaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()));
            }

            //[Content_Types].xml
            {
                var sb = new StringBuilder(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"" Extension=""bin""/><Default ContentType=""application/xml"" Extension=""xml""/><Default ContentType=""image/jpeg"" Extension=""jpg""/><Default ContentType=""image/png"" Extension=""png""/><Default ContentType=""image/gif"" Extension=""gif""/><Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>");
                foreach (var p in _zipDictionary)
                    sb.Append($"<Override ContentType=\"{p.Value.ContentType}\" PartName=\"/{p.Key}\" />");
                sb.Append("</Types>");
                ZipArchiveEntry entry = _archive.CreateEntry("[Content_Types].xml", CompressionLevel.Fastest);
                using (var zipStream = entry.Open())
                using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
                    writer.Write(sb.ToString());
            }
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

        public void Insert()
        {
            throw new NotImplementedException();
        }
    }
}
