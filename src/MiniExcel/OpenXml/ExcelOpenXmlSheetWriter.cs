using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
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
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly MiniExcelZipArchive _archive;
        private readonly static UTF8Encoding _utf8WithBom = new UTF8Encoding(true);
        private readonly OpenXmlConfiguration _configuration;
        private readonly Stream _stream;
        private readonly bool _printHeader;
        private readonly object _value;
        private readonly string _defaultSheetName;
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
            this._defaultSheetName = sheetName;
        }

        public ExcelOpenXmlSheetWriter()
        {
        }

        public void SaveAs()
        {
            GenerateDefaultOpenXml();

            var sheets = GetSheets();

            foreach (var sheet in sheets)
            {
                _sheets.Add(sheet.Item1); //TODO:remove
                currentSheetIndex = sheet.Item1.SheetIdx;
                CreateSheetXml(sheet.Item2, sheet.Item1.Path);
            }

            GenerateEndXml();
            _archive.Dispose();
        }

        internal void GenerateDefaultOpenXml()
        {
            CreateZipEntry(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels);
            CreateZipEntry(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString);
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
                }
                else
                {
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
            }
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
        }

        private void WriteEmptySheet(MiniExcelStreamWriter writer)
        {
            writer.Write(ExcelXml.EmptySheetXml);
        }

        private void GenerateSheetByIDataReader(MiniExcelStreamWriter writer, IDataReader reader)
        {
            long dimensionWritePosition = 0;
            writer.Write(WorksheetXml.StartWorksheet);
            var yIndex = 1;
            int maxColumnIndex;
            int maxRowIndex;
            {

                if (_configuration.FastMode)
                {
                    dimensionWritePosition = writer.WriteAndFlush(WorksheetXml.StartDimension);
                    writer.Write(WorksheetXml.DimensionPlaceholder); // end of code will be replaced
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

                writer.Write(WorksheetXml.StartSheetData);
                int fieldCount = reader.FieldCount;
                if (_printHeader)
                {
                    PrintHeader(writer, props);
                    yIndex++;
                }

                while (reader.Read())
                {
                    writer.Write(WorksheetXml.StartRow(yIndex));
                    var xIndex = 1;
                    for (int i = 0; i < fieldCount; i++)
                    {
                        var cellValue = reader.GetValue(i);
                        WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: null);
                        xIndex++;
                    }
                    writer.Write(WorksheetXml.EndRow);
                    yIndex++;
                }

                // Subtract 1 because cell indexing starts with 1
                maxRowIndex = yIndex - 1;
            }
            writer.Write(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                writer.Write(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            writer.WriteAndFlush(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode)
            {
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
            }
        }

        private void GenerateSheetByEnumerable(MiniExcelStreamWriter writer, IEnumerable values)
        {
            var maxRowIndex = 0;
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

            int maxColumnIndex;
            List<ExcelColumnInfo> props;
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
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(genericDic, null, _configuration);
                    maxColumnIndex = props.Count;
                }
                else if (firstItem is IDictionary dic)
                {
                    mode = "IDictionary";
                    props = CustomPropertyHelper.GetDictionaryColumnInfo(null, dic, _configuration);
                    //maxColumnIndex = dic.Keys.Count;
                    maxColumnIndex = props.Count; // why not using keys, because ignore attribute ![image](https://user-images.githubusercontent.com/12729184/163686902-286abb70-877b-4e84-bd3b-001ad339a84a.png)
                }
                else
                {
                    SetGenericTypePropertiesMode(firstItem.GetType(), ref mode, out maxColumnIndex, out props);
                }
            }

            writer.Write(WorksheetXml.StartWorksheetWithRelationship);

            long dimensionWritePosition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && rowCount == null)
            {
                // Write a placeholder for the table dimensions and save thee position for later
                dimensionWritePosition = writer.WriteAndFlush(WorksheetXml.StartDimension);
                writer.Write(WorksheetXml.DimensionPlaceholder);
            }
            else
            {
                maxRowIndex = rowCount.Value + (_printHeader && rowCount > 0 ? 1 : 0);
                writer.Write(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            //cols:width
            WriteColumnsWidths(writer, props);

            //header
            writer.Write(WorksheetXml.StartSheetData);
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
                switch (mode)
                {
                    case "IDictionary<string, object>": //Dapper Row
                        maxRowIndex = GenerateSheetByColumnInfo<IDictionary<string, object>>(writer, enumerator, props, xIndex, yIndex);
                        break;
                    case "IDictionary":
                        maxRowIndex = GenerateSheetByColumnInfo<IDictionary>(writer, enumerator, props, xIndex, yIndex);
                        break;
                    case "Properties":
                        maxRowIndex = GenerateSheetByColumnInfo<object>(writer, enumerator, props, xIndex, yIndex);
                        break;
                    default:
                        throw new NotImplementedException($"Type {values.GetType().FullName} is not implemented. Please open an issue.");
                }
            }

            writer.Write(WorksheetXml.EndSheetData);
            if (_configuration.AutoFilter)
            {
                writer.Write(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            writer.Write(WorksheetXml.Drawing(currentSheetIndex));
            writer.Write(WorksheetXml.EndWorksheet);

            // The dimension has already been written if row count is defined
            if (_configuration.FastMode && rowCount == null)
            {
                // Seek back and write the dimensions of the table
                writer.SetPosition(dimensionWritePosition);
                writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");
            }
        }

        private void GenerateSheetByDataTable(MiniExcelStreamWriter writer, DataTable value)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write(WorksheetXml.StartWorksheet);

            var yIndex = xy.Item2;

            // dimension
            var maxRowIndex = value.Rows.Count + (_printHeader && value.Rows.Count > 0 ? 1 : 0);
            var maxColumnIndex = value.Columns.Count;
            writer.Write(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, maxColumnIndex)));

            var props = new List<ExcelColumnInfo>();
            for (var i = 0; i < value.Columns.Count; i++)
            {
                var columnName = value.Columns[i].Caption ?? value.Columns[i].ColumnName;
                var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                props.Add(prop);
            }

            WriteColumnsWidths(writer, props);

            writer.Write(WorksheetXml.StartSheetData);
            if (_printHeader)
            {
                writer.Write(WorksheetXml.StartRow(yIndex));
                var xIndex = xy.Item1;
                foreach (var p in props)
                {
                    var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                    WriteCell(writer, r, columnName: p.ExcelColumnName);
                    xIndex++;
                }

                writer.Write(WorksheetXml.EndRow);
                yIndex++;
            }

            for (int i = 0; i < value.Rows.Count; i++)
            {
                writer.Write(WorksheetXml.StartRow(yIndex));
                var xIndex = xy.Item1;

                for (int j = 0; j < value.Columns.Count; j++)
                {
                    var cellValue = value.Rows[i][j];
                    WriteCell(writer, yIndex, xIndex, cellValue, columnInfo: null);
                    xIndex++;
                }
                writer.Write(WorksheetXml.EndRow);
                yIndex++;
            }

            writer.Write(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                writer.Write(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            writer.Write(WorksheetXml.EndWorksheet);
        }

        private static void WriteColumnsWidths(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnInfo> props)
        {
            var ecwProps = props.Where(x => x?.ExcelColumnWidth != null).ToList();
            if (ecwProps.Count <= 0)
                return;
            writer.Write(WorksheetXml.StartCols);
            foreach (var p in ecwProps)
            {
                writer.Write(WorksheetXml.Column(p.ExcelColumnIndex, p.ExcelColumnWidth));
            }

            writer.Write(WorksheetXml.EndCols);
        }

        private static void PrintHeader(MiniExcelStreamWriter writer, List<ExcelColumnInfo> props)
        {
            var xIndex = 1;
            var yIndex = 1;
            writer.Write(WorksheetXml.StartRow(yIndex));

            foreach (var p in props)
            {
                if (p == null)
                {
                    xIndex++; //reason : https://github.com/shps951023/MiniExcel/issues/142
                    continue;
                }

                var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                WriteCell(writer, r, columnName: p.ExcelColumnName);
                xIndex++;
            }

            writer.Write(WorksheetXml.EndRow);
        }

        private int GenerateSheetByColumnInfo<T>(MiniExcelStreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, int xIndex = 1, int yIndex = 1)
        {
            var isDic = typeof(T) == typeof(IDictionary);
            var isDapperRow = typeof(T) == typeof(IDictionary<string, object>);
            do
            {
                // The enumerator has already moved to the first item
                T v = (T)value.Current;

                writer.Write(WorksheetXml.StartRow(yIndex));
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

                writer.Write(WorksheetXml.EndRow);
                yIndex++;
            } while (value.MoveNext());

            return yIndex - 1;
        }

        private void WriteCell(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo)
        {
            var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                writer.Write(WorksheetXml.EmptyCell(columnReference, "2"));
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

            var styleIndex = tuple.Item1; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
            var dataType = tuple.Item2; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1
            var cellValue = tuple.Item3;

            /*Prefix and suffix blank space will lost after SaveAs #294*/
            var preserveSpace = cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) || cellValue.EndsWith(" ", StringComparison.Ordinal));
            writer.Write(WorksheetXml.Cell(columnReference, dataType, styleIndex, cellValue, preserveSpace: preserveSpace));
        }

        private static void WriteCell(MiniExcelStreamWriter writer, string cellReference, string columnName)
            => writer.Write(WorksheetXml.Cell(cellReference, "str", "1", ExcelOpenXmlUtils.EncodeXML(columnName)));

        private void GenerateEndXml()
        {
            AddFilesToZip();

            GenerateStylesXml();

            GenerateDrawinRelXml();

            GenerateDrawingXml();

            GenerateWorkbookXml();

            GenerateContentTypesXml();
        }

        private void AddFilesToZip()
        {
            foreach (var item in _files)
            {
                this.CreateZipEntry(item.Path, item.Byte);
            }
        }

        /// <summary>
        /// styles.xml
        /// </summary>
        private void GenerateStylesXml()
        {
            var styleXml = GetStylesXml();
            CreateZipEntry(ExcelFileNames.Styles, ExcelContentTypes.Styles, styleXml);
        }

        private void GenerateDrawinRelXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingRelationshipXml(sheetIndex);
                CreateZipEntry(
                    ExcelFileNames.DrawingRels(sheetIndex),
                    null,
                    ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing));
            }
        }

        private void GenerateDrawingXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingXml(sheetIndex);

                CreateZipEntry(
                    ExcelFileNames.Drawing(sheetIndex),
                    ExcelContentTypes.Drawing,
                    ExcelXml.DefaultDrawing.Replace("{{format}}", drawing));
            }
        }

        /// <summary>
        /// workbook.xml、workbookRelsXml
        /// </summary>
        private void GenerateWorkbookXml()
        {
            GenerateWorkBookXmls(
                out StringBuilder workbookXml,
                out StringBuilder workbookRelsXml,
                out Dictionary<int, string> sheetsRelsXml);

            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                CreateZipEntry(
                    ExcelFileNames.SheetRels(sheetRelsXml.Key),
                    null,
                    ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value));
            }

            CreateZipEntry(
                ExcelFileNames.Workbook,
                ExcelContentTypes.Workbook,
                ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()));

            CreateZipEntry(
                ExcelFileNames.WorkbookRels,
                null,
                ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()));
        }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        private void GenerateContentTypesXml()
        {
            var contentTypes = GetContentTypesXml();

            CreateZipEntry(ExcelFileNames.ContentTypes, null, contentTypes);
        }

        private void CreateZipEntry(string path, string contentType, string content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
            {
                writer.Write(content);
            }

            if (!string.IsNullOrEmpty(contentType))
            {
                _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
            }
        }

        private void CreateZipEntry(string path, byte[] content)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            {
                zipStream.Write(content, 0, content.Length);
            }
        }
    }
}