using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.OpenXml.Styles;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

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
            if (_configuration.EnableAutoWidth && !_configuration.FastMode)
            {
                throw new InvalidOperationException("Auto width requires fast mode to be enabled");
            }

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

        public void Insert(bool overwriteSheet = false)
        {
            if (!_configuration.FastMode)
            {
                throw new InvalidOperationException("Insert requires fast mode to be enabled");
            }

            var sheetRecords = new ExcelOpenXmlSheetReader(_stream, _configuration).GetWorkbookRels(_archive.Entries).ToArray();
            foreach (var sheetRecord in sheetRecords.OrderBy(o => o.Id))
            {
                _sheets.Add(new SheetDto { Name = sheetRecord.Name, SheetIdx = (int)sheetRecord.Id, State = sheetRecord.State });
            }
            var existSheetDto = _sheets.SingleOrDefault(s => s.Name == _defaultSheetName);
            if (existSheetDto != null && !overwriteSheet)
            {
                throw new Exception($"Sheet “{_defaultSheetName}” already exist");
            }

            GenerateStylesXml();//GenerateStylesXml必须在校验overwriteSheet之后，避免不必要的样式更改

            if (existSheetDto == null)
            {
                currentSheetIndex = (int)sheetRecords.Max(m => m.Id) + 1;
                var insertSheetInfo = GetSheetInfos(_defaultSheetName);
                var insertSheetDto = insertSheetInfo.ToDto(currentSheetIndex);
                _sheets.Add(insertSheetDto);
                CreateSheetXml(_value, insertSheetDto.Path);
            }
            else
            {
                currentSheetIndex = existSheetDto.SheetIdx;
                _archive.Entries.Single(s => s.FullName == existSheetDto.Path).Delete();
                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.DrawingRels(currentSheetIndex))?.Delete();
                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Drawing(currentSheetIndex))?.Delete();
                CreateSheetXml(_value, existSheetDto.Path);
            }

            AddFilesToZip();

            GenerateDrawinRelXml(currentSheetIndex);

            GenerateDrawingXml(currentSheetIndex);

            GenerateWorkBookXmls(out StringBuilder workbookXml, out StringBuilder workbookRelsXml, out Dictionary<int, string> sheetsRelsXml);

            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                var sheetRelsXmlPath = ExcelFileNames.SheetRels(sheetRelsXml.Key);
                _archive.Entries.SingleOrDefault(s => s.FullName == sheetRelsXmlPath)?.Delete();
                CreateZipEntry(sheetRelsXmlPath, null, ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value));
            }

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Workbook)?.Delete();
            CreateZipEntry(ExcelFileNames.Workbook, ExcelContentTypes.Workbook, ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()));

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.WorkbookRels)?.Delete();
            CreateZipEntry(ExcelFileNames.WorkbookRels, null, ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()));

            _archive.Dispose();
        }

        internal void GenerateDefaultOpenXml()
        {
            CreateZipEntry(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels);
            CreateZipEntry(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString);
            GenerateStylesXml();
        }

        private void CreateSheetXml(object values, string sheetPath)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelStreamWriter writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
            {
                if (values == null)
                {
                    WriteEmptySheet(writer);
                }
                else
                {
                    WriteValues(writer, values);
                }
            }
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
        }

        private void WriteEmptySheet(MiniExcelStreamWriter writer)
        {
            writer.Write(ExcelXml.EmptySheetXml);
        }

        private long WriteDimensionPlaceholder(MiniExcelStreamWriter writer)
        {
            var dimensionPlaceholderPostition = writer.WriteAndFlush(WorksheetXml.StartDimension);
            writer.Write(WorksheetXml.DimensionPlaceholder); // end of code will be replaced

            return dimensionPlaceholderPostition;
        }

        private void WriteDimension(MiniExcelStreamWriter writer, int maxRowIndex, int maxColumnIndex, long placeholderPosition)
        {
            // Flush and save position so that we can get back again.
            var position = writer.Flush();

            writer.SetPosition(placeholderPosition);
            writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");

            writer.SetPosition(position);
        }


        private void WriteValues(MiniExcelStreamWriter writer, object values)
        {
            IMiniExcelWriteAdapter writeAdapter;
            switch (values)
            {
                case IDataReader dataReader:
                    writeAdapter = new DataReaderWriteAdapter(dataReader, _configuration);
                    break;
                case IEnumerable enumerable:
                    writeAdapter = new EnumerableWriteAdapter(enumerable, _configuration);
                    break;
                case DataTable dataTable:
                    writeAdapter = new DataTableWriteAdapter(dataTable, _configuration);
                    break;
                default:
                    throw new NotImplementedException();
            }

            var hasCount = writeAdapter.TryGetNonEnumeratedCount(out var count);
            var props = writeAdapter.GetColumns();
            var maxColumnIndex = props.Count;
            int maxRowIndex;
            if (props.Count == 0)
            {
                WriteEmptySheet(writer);
                return;
            }

            writer.Write(WorksheetXml.StartWorksheetWithRelationship);

            long dimensionPlaceholderPostition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && !hasCount)
            {
                dimensionPlaceholderPostition = WriteDimensionPlaceholder(writer);
            }
            else
            {
                maxRowIndex = count + (_printHeader && count > 0 ? 1 : 0);
                writer.Write(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, props.Count)));
            }

            //sheet view
            writer.Write(GetSheetViews());

            //cols:width
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = WriteColumnWidthPlaceholders(writer, props);
                widths = new ExcelWidthCollection(_configuration.MinWidth, _configuration.MaxWidth, props);
            }
            else
            {
                WriteColumnsWidths(writer, ExcelColumnWidth.FromProps(props));
            }

            //header
            writer.Write(WorksheetXml.StartSheetData);
            var currentRowIndex = 0;
            if (_printHeader)
            {
                PrintHeader(writer, props);
                currentRowIndex++;
            }

            foreach (var row in writeAdapter.GetRows(props))
            {
                writer.Write(WorksheetXml.StartRow(++currentRowIndex));
                foreach (var cellValue in row)
                {
                    WriteCell(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Prop, widths);
                }
                writer.Write(WorksheetXml.EndRow);

                ;
            }
            maxRowIndex = currentRowIndex;

            writer.Write(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                writer.Write(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            writer.Write(WorksheetXml.Drawing(currentSheetIndex));
            writer.Write(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode && dimensionPlaceholderPostition != default)
            {
                WriteDimension(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }
            if (_configuration.EnableAutoWidth)
            {
                OverWriteColumnWidthPlaceholders(writer, columnWidthsPlaceholderPosition, widths.Columns);
            }
        }


        private long WriteColumnWidthPlaceholders(MiniExcelStreamWriter writer, ICollection<ExcelColumnInfo> props)
        {
            var placeholderPosition = writer.Flush();
            writer.WriteWhitespace(WorksheetXml.GetColumnPlaceholderLength(props.Count));
            return placeholderPosition;
        }

        private void OverWriteColumnWidthPlaceholders(MiniExcelStreamWriter writer, long placeholderPosition, IEnumerable<ExcelColumnWidth> columnWidths)
        {
            var position = writer.Flush();

            writer.SetPosition(placeholderPosition);
            WriteColumnsWidths(writer, columnWidths);

            writer.Flush();
            writer.SetPosition(position);
        }

        private void WriteColumnsWidths(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnWidth> columnWidths)
        {
            var hasWrittenStart = false;
            foreach (var column in columnWidths)
            {
                if (!hasWrittenStart)
                {
                    writer.Write(WorksheetXml.StartCols);
                    hasWrittenStart = true;
                }
                writer.Write(WorksheetXml.Column(column.Index, column.Width));
            }
            if (!hasWrittenStart)
            {
                return;
            }
            writer.Write(WorksheetXml.EndCols);
        }

        private void PrintHeader(MiniExcelStreamWriter writer, List<ExcelColumnInfo> props)
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

        private int GenerateSheetByColumnInfo<T>(MiniExcelStreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, ExcelWidthCollection widthCollection, int xIndex = 1, int yIndex = 1)
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

                    WriteCell(writer, yIndex, cellIndex, cellValue, columnInfo, widthCollection);

                    cellIndex++;
                }

                writer.Write(WorksheetXml.EndRow);
                yIndex++;
            } while (value.MoveNext());

            return yIndex - 1;
        }

        private void WriteCell(MiniExcelStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo columnInfo, ExcelWidthCollection widthCollection)
        {
            var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                writer.Write(WorksheetXml.EmptyCell(columnReference, GetCellXfId("2")));
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, columnInfo, valueIsNull);

            var styleIndex = tuple.Item1; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cell?view=openxml-3.0.1
            var dataType = tuple.Item2; // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalues?view=openxml-3.0.1
            var cellValue = tuple.Item3;

            if (columnInfo?.CustomFormatter != null)
            {
                try
                {
                    cellValue = columnInfo.CustomFormatter(cellValue);
                }
                catch
                {
                    //ignored
                }
            }

            var columnType = columnInfo?.ExcelColumnType ?? ColumnType.Value;

            /*Prefix and suffix blank space will lost after SaveAs #294*/
            var preserveSpace = cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) || cellValue.EndsWith(" ", StringComparison.Ordinal));
            writer.Write(WorksheetXml.Cell(columnReference, dataType, GetCellXfId(styleIndex), cellValue, preserveSpace: preserveSpace, columnType: columnType));
            widthCollection?.AdjustWidth(cellIndex, cellValue);
        }

        private void WriteCell(MiniExcelStreamWriter writer, string cellReference, string columnName)
            => writer.Write(WorksheetXml.Cell(cellReference, "str", GetCellXfId("1"), ExcelOpenXmlUtils.EncodeXML(columnName)));

        private void GenerateEndXml()
        {
            AddFilesToZip();

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
            using (var context = new SheetStyleBuildContext(_zipDictionary, _archive, _utf8WithBom, _configuration.DynamicColumns))
            {
                var builder = (ISheetStyleBuilder)null;
                switch (_configuration.TableStyles)
                {
                    case TableStyles.None:
                        builder = new MinimalSheetStyleBuilder(context);
                        break;
                    case TableStyles.Default:
                        builder = new DefaultSheetStyleBuilder(context);
                        break;
                }
                var result = builder.Build();
                cellXfIdMap = result.CellXfIdMap;
            }
        }

        private void GenerateDrawinRelXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                GenerateDrawinRelXml(sheetIndex);
            }
        }

        private void GenerateDrawinRelXml(int sheetIndex)
        {
            var drawing = GetDrawingRelationshipXml(sheetIndex);
            CreateZipEntry(
                ExcelFileNames.DrawingRels(sheetIndex),
                null,
                ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing));
        }

        private void GenerateDrawingXml()
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                GenerateDrawingXml(sheetIndex);
            }
        }

        private void GenerateDrawingXml(int sheetIndex)
        {
            var drawing = GetDrawingXml(sheetIndex);
            CreateZipEntry(
                ExcelFileNames.Drawing(sheetIndex),
                ExcelContentTypes.Drawing,
                ExcelXml.DefaultDrawing.Replace("{{format}}", drawing));
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
