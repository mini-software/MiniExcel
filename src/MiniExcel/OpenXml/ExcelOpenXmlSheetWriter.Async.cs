using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.OpenXml.Styles;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        public async Task SaveAsAsync(CancellationToken cancellationToken = default)
        {
            await GenerateDefaultOpenXmlAsync(cancellationToken);

            var sheets = GetSheets();

            foreach (var sheet in sheets)
            {
                _sheets.Add(sheet.Item1); //TODO:remove
                currentSheetIndex = sheet.Item1.SheetIdx;
                await CreateSheetXmlAsync(sheet.Item2, sheet.Item1.Path, cancellationToken);
            }

            await GenerateEndXmlAsync(cancellationToken);
            _archive.Dispose();
        }

        public async Task InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default)
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

            await GenerateStylesXmlAsync(cancellationToken);//GenerateStylesXml必须在校验overwriteSheet之后，避免不必要的样式更改

            if (existSheetDto == null)
            {
                currentSheetIndex = (int)sheetRecords.Max(m => m.Id) + 1;
                var insertSheetInfo = GetSheetInfos(_defaultSheetName);
                var insertSheetDto = insertSheetInfo.ToDto(currentSheetIndex);
                _sheets.Add(insertSheetDto);
                await CreateSheetXmlAsync(_value, insertSheetDto.Path, cancellationToken);
            }
            else
            {
                currentSheetIndex = existSheetDto.SheetIdx;
                _archive.Entries.Single(s => s.FullName == existSheetDto.Path).Delete();
                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.DrawingRels(currentSheetIndex))?.Delete();
                _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Drawing(currentSheetIndex))?.Delete();
                await CreateSheetXmlAsync(_value, existSheetDto.Path, cancellationToken);
            }

            await AddFilesToZipAsync(cancellationToken);

            await GenerateDrawinRelXmlAsync(currentSheetIndex, cancellationToken);

            await GenerateDrawingXmlAsync(currentSheetIndex, cancellationToken);

            GenerateWorkBookXmls(out StringBuilder workbookXml, out StringBuilder workbookRelsXml, out Dictionary<int, string> sheetsRelsXml);

            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                var sheetRelsXmlPath = ExcelFileNames.SheetRels(sheetRelsXml.Key);
                _archive.Entries.SingleOrDefault(s => s.FullName == sheetRelsXmlPath)?.Delete();
                await CreateZipEntryAsync(sheetRelsXmlPath, null, ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value), cancellationToken);
            }

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Workbook)?.Delete();
            await CreateZipEntryAsync(ExcelFileNames.Workbook, ExcelContentTypes.Workbook, ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()), cancellationToken);

            _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.WorkbookRels)?.Delete();
            await CreateZipEntryAsync(ExcelFileNames.WorkbookRels, null, ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()), cancellationToken);

            _archive.Dispose();
        }

        internal async Task GenerateDefaultOpenXmlAsync(CancellationToken cancellationToken)
        {
            await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken);
            await CreateZipEntryAsync(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString, cancellationToken);
            await GenerateStylesXmlAsync(cancellationToken);
        }

        private async Task CreateSheetXmlAsync(object values, string sheetPath, CancellationToken cancellationToken)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelAsyncStreamWriter writer = new MiniExcelAsyncStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize, cancellationToken))
            {
                if (values == null)
                {
                    await WriteEmptySheetAsync(writer);
                }
                else
                {
                    await WriteValuesAsync(writer, values, cancellationToken);
                }
            }
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
        }

        private async Task WriteEmptySheetAsync(MiniExcelAsyncStreamWriter writer)
        {
            await writer.WriteAsync(ExcelXml.EmptySheetXml);
        }

        private async Task<long> WriteDimensionPlaceholderAsync(MiniExcelAsyncStreamWriter writer)
        {
            var dimensionPlaceholderPostition = await writer.WriteAndFlushAsync(WorksheetXml.StartDimension);
            await writer.WriteAsync(WorksheetXml.DimensionPlaceholder); // end of code will be replaced

            return dimensionPlaceholderPostition;
        }

        private async Task WriteDimensionAsync(MiniExcelAsyncStreamWriter writer, int maxRowIndex, int maxColumnIndex, long placeholderPosition)
        {
            // Flush and save position so that we can get back again.
            var position = await writer.FlushAsync();

            writer.SetPosition(placeholderPosition);
            await writer.WriteAndFlushAsync($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");

            writer.SetPosition(position);
        }

        private async Task WriteValuesAsync(MiniExcelAsyncStreamWriter writer, object values, CancellationToken cancellationToken)
        {
#if NETSTANDARD2_0_OR_GREATER || NET
            IMiniExcelWriteAdapter writeAdapter = null;
            if (!MiniExcelWriteAdapterFactory.TryGetAsyncWriteAdapter(values, _configuration, out var asyncWriteAdapter))
            {
                writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);
            }

            var count = 0;
            var isKnownCount = writeAdapter != null && writeAdapter.TryGetKnownCount(out count);
            var props = writeAdapter != null ? writeAdapter?.GetColumns() : await asyncWriteAdapter.GetColumnsAsync();
#else
            IMiniExcelWriteAdapter writeAdapter =  MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);

            var isKnownCount = writeAdapter.TryGetKnownCount(out var count);
            var props = writeAdapter.GetColumns();
#endif

            if (props == null)
            {
                await WriteEmptySheetAsync(writer);
                return;
            }
            var maxColumnIndex = props.Count;
            int maxRowIndex;

            await writer.WriteAsync(WorksheetXml.StartWorksheetWithRelationship);

            long dimensionPlaceholderPostition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && !isKnownCount)
            {
                dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer);
            }
            else if (isKnownCount)
            {
                maxRowIndex = count + (_printHeader ? 1 : 0);
                await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, props.Count)));
            }

            //sheet view
            await writer.WriteAsync(GetSheetViews());

            //cols:width
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = await WriteColumnWidthPlaceholdersAsync(writer, props);
                widths = new ExcelWidthCollection(_configuration.MinWidth, _configuration.MaxWidth, props);
            }
            else
            {
                await WriteColumnsWidthsAsync(writer, ExcelColumnWidth.FromProps(props));
            }

            //header
            await writer.WriteAsync(WorksheetXml.StartSheetData);
            var currentRowIndex = 0;
            if (_printHeader)
            {
                await PrintHeaderAsync(writer, props);
                currentRowIndex++;
            }

            if (writeAdapter != null)
            {
                foreach (var row in writeAdapter.GetRows(props, cancellationToken))
                {
                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex));
                    foreach (var cellValue in row)
                    {
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Prop, widths);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow);
                }
            }
#if NETSTANDARD2_0_OR_GREATER || NET
            else
            {
                await foreach (var row in asyncWriteAdapter.GetRowsAsync(props, cancellationToken))
                {
                    await writer.WriteAsync(WorksheetXml.StartRow(++currentRowIndex));
                    await foreach (var cellValue in row)
                    {
                        await WriteCellAsync(writer, currentRowIndex, cellValue.CellIndex, cellValue.Value, cellValue.Prop, widths);
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow);
                }
            }
#endif

            maxRowIndex = currentRowIndex;

            await writer.WriteAsync(WorksheetXml.Drawing(currentSheetIndex));
            await writer.WriteAsync(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            await writer.WriteAsync(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode && dimensionPlaceholderPostition != default)
            {
                await WriteDimensionAsync(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }
            if (_configuration.EnableAutoWidth)
            {
                await OverWriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths.Columns);
            }
        }

        private async Task<long> WriteColumnWidthPlaceholdersAsync(MiniExcelAsyncStreamWriter writer, ICollection<ExcelColumnInfo> props)
        {
            var placeholderPosition = await writer.FlushAsync();
            await writer.WriteWhitespaceAsync(WorksheetXml.GetColumnPlaceholderLength(props.Count));
            return placeholderPosition;
        }

        private async Task OverWriteColumnWidthPlaceholdersAsync(MiniExcelAsyncStreamWriter writer, long placeholderPosition, IEnumerable<ExcelColumnWidth> columnWidths)
        {
            var position = await writer.FlushAsync();

            writer.SetPosition(placeholderPosition);
            await WriteColumnsWidthsAsync(writer, columnWidths);

            await writer.FlushAsync();
            writer.SetPosition(position);
        }

        private async Task WriteColumnsWidthsAsync(MiniExcelAsyncStreamWriter writer, IEnumerable<ExcelColumnWidth> columnWidths)
        {
            var hasWrittenStart = false;
            foreach (var column in columnWidths)
            {
                if (!hasWrittenStart)
                {
                    await writer.WriteAsync(WorksheetXml.StartCols);
                    hasWrittenStart = true;
                }
                await writer.WriteAsync(WorksheetXml.Column(column.Index, column.Width));
            }
            if (!hasWrittenStart)
            {
                return;
            }
            await writer.WriteAsync(WorksheetXml.EndCols);
        }

        private async Task PrintHeaderAsync(MiniExcelAsyncStreamWriter writer, List<ExcelColumnInfo> props)
        {
            var xIndex = 1;
            var yIndex = 1;
            await writer.WriteAsync(WorksheetXml.StartRow(yIndex));

            foreach (var p in props)
            {
                if (p == null)
                {
                    xIndex++; //reason : https://github.com/shps951023/MiniExcel/issues/142
                    continue;
                }

                var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                await WriteCellAsync(writer, r, columnName: p.ExcelColumnName);
                xIndex++;
            }

            await writer.WriteAsync(WorksheetXml.EndRow);
        }

        private async Task WriteCellAsync(MiniExcelAsyncStreamWriter writer, string cellReference, string columnName)
        {
            await writer.WriteAsync(WorksheetXml.Cell(cellReference, "str", GetCellXfId("1"), ExcelOpenXmlUtils.EncodeXML(columnName)));
        }

        private async Task WriteCellAsync(MiniExcelAsyncStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo p, ExcelWidthCollection widthCollection)
        {
            var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, GetCellXfId("2")));
                return;
            }

            if (p.CustomFormatter != null)
            {
                try
                {
                    value = p.CustomFormatter(value);
                }
                catch
                {
                    //ignored
                }
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, p, valueIsNull);

            var styleIndex = tuple.Item1;
            var dataType = tuple.Item2;
            var cellValue = tuple.Item3;
            var columnType = p.ExcelColumnType;

            /*Prefix and suffix blank space will lost after SaveAs #294*/
            var preserveSpace = cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) ||
                                                      cellValue.EndsWith(" ", StringComparison.Ordinal));

            await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, GetCellXfId(styleIndex), cellValue, preserveSpace: preserveSpace, columnType: columnType));
            widthCollection?.AdjustWidth(cellIndex, cellValue);
        }

        private async Task GenerateEndXmlAsync(CancellationToken cancellationToken)
        {
            await AddFilesToZipAsync(cancellationToken);

            await GenerateDrawinRelXmlAsync(cancellationToken);

            await GenerateDrawingXmlAsync(cancellationToken);

            await GenerateWorkbookXmlAsync(cancellationToken);

            await GenerateContentTypesXmlAsync(cancellationToken);
        }

        private async Task AddFilesToZipAsync(CancellationToken cancellationToken)
        {
            foreach (var item in _files)
            {
                await this.CreateZipEntryAsync(item.Path, item.Byte, cancellationToken);
            }
        }

        /// <summary>
        /// styles.xml
        /// </summary>
        private async Task GenerateStylesXmlAsync(CancellationToken cancellationToken)
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
                var result = await builder.BuildAsync(cancellationToken);
                cellXfIdMap = result.CellXfIdMap;
            }
        }

        private async Task GenerateDrawinRelXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                await GenerateDrawinRelXmlAsync(sheetIndex, cancellationToken);
            }
        }

        private async Task GenerateDrawinRelXmlAsync(int sheetIndex, CancellationToken cancellationToken)
        {
            var drawing = GetDrawingRelationshipXml(sheetIndex);
            await CreateZipEntryAsync(
                ExcelFileNames.DrawingRels(sheetIndex),
                string.Empty,
                ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing),
                cancellationToken);
        }

        private async Task GenerateDrawingXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                await GenerateDrawingXmlAsync(sheetIndex, cancellationToken);
            }
        }

        private async Task GenerateDrawingXmlAsync(int sheetIndex, CancellationToken cancellationToken)
        {
            var drawing = GetDrawingXml(sheetIndex);
            await CreateZipEntryAsync(
                ExcelFileNames.Drawing(sheetIndex),
                ExcelContentTypes.Drawing,
                ExcelXml.DefaultDrawing.Replace("{{format}}", drawing),
                cancellationToken);
        }

        /// <summary>
        /// workbook.xml 、 workbookRelsXml
        /// </summary>
        private async Task GenerateWorkbookXmlAsync(CancellationToken cancellationToken)
        {
            GenerateWorkBookXmls(
                out StringBuilder workbookXml,
                out StringBuilder workbookRelsXml,
                out Dictionary<int, string> sheetsRelsXml);

            foreach (var sheetRelsXml in sheetsRelsXml)
            {
                await CreateZipEntryAsync(
                    ExcelFileNames.SheetRels(sheetRelsXml.Key),
                    null,
                    ExcelXml.DefaultSheetRelXml.Replace("{{format}}", sheetRelsXml.Value),
                    cancellationToken);
            }

            await CreateZipEntryAsync(
                ExcelFileNames.Workbook,
                ExcelContentTypes.Workbook,
                ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml.ToString()),
                    cancellationToken);

            await CreateZipEntryAsync(
                ExcelFileNames.WorkbookRels,
                null,
                ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml.ToString()),
                    cancellationToken);
        }

        /// <summary>
        /// [Content_Types].xml
        /// </summary>
        private async Task GenerateContentTypesXmlAsync(CancellationToken cancellationToken)
        {
            var contentTypes = GetContentTypesXml();

            await CreateZipEntryAsync(ExcelFileNames.ContentTypes, null, contentTypes, cancellationToken);
        }

        private async Task CreateZipEntryAsync(string path, string contentType, string content, CancellationToken cancellationToken)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelAsyncStreamWriter writer = new MiniExcelAsyncStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize, cancellationToken))
                await writer.WriteAsync(content);
            if (!string.IsNullOrEmpty(contentType))
                _zipDictionary.Add(path, new ZipPackageInfo(entry, contentType));
        }

        private async Task CreateZipEntryAsync(string path, byte[] content, CancellationToken cancellationToken)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(path, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
                await zipStream.WriteAsync(content, 0, content.Length, cancellationToken);
        }
    }
}