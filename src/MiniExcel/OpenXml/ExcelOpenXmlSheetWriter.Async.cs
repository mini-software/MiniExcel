using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        public async Task SaveAsAsync(CancellationToken cancellationToken = default(CancellationToken))
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

        internal async Task GenerateDefaultOpenXmlAsync(CancellationToken cancellationToken)
        {
            await CreateZipEntryAsync(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels, cancellationToken);
            await CreateZipEntryAsync(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString, cancellationToken);
            await GenerateStylesXmlAsync(cancellationToken);
        }

        private async Task CreateSheetXmlAsync(object value, string sheetPath, CancellationToken cancellationToken)
        {
            ZipArchiveEntry entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            using (var zipStream = entry.Open())
            using (MiniExcelAsyncStreamWriter writer = new MiniExcelAsyncStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize, cancellationToken))
            {
                if (value == null)
                {
                    await WriteEmptySheetAsync(writer);
                }
                else
                {
                    //DapperRow

                    switch (value)
                    {
                        case IDataReader dataReader:
                            await GenerateSheetByIDataReaderAsync(writer, dataReader);
                            break;
                        case IEnumerable enumerable:
                            await GenerateSheetByEnumerableAsync(writer, enumerable);
                            break;
                        case DataTable dataTable:
                            await GenerateSheetByDataTableAsync(writer, dataTable);
                            break;
                        default:
                            throw new NotImplementedException($"Type {value.GetType().FullName} is not implemented. Please open an issue.");
                    }
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

        private async Task GenerateSheetByIDataReaderAsync(MiniExcelAsyncStreamWriter writer, IDataReader reader)
        {
            long dimensionPlaceholderPostition = 0;
            await writer.WriteAsync(WorksheetXml.StartWorksheet);
            var yIndex = 1;
            int maxColumnIndex;
            int maxRowIndex;
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            {
                if (_configuration.FastMode)
                {
                    dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer);
                }

                var props = new List<ExcelColumnInfo>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);

                    if (!_configuration.DynamicColumnFirst)
                    {
                        var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                        props.Add(prop);
                        continue;
                    }

                    if (_configuration
                        .DynamicColumns
                        .Any(a => string.Equals(
                            a.Key,
                            columnName,
                            StringComparison.OrdinalIgnoreCase)))

                    {
                        var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                        props.Add(prop);
                    }
                }
                maxColumnIndex = props.Count;

                //sheet view
                await writer.WriteAsync(GetSheetViews());

                if (_configuration.EnableAutoWidth)
                {
                    columnWidthsPlaceholderPosition = await WriteColumnWidthPlaceholdersAsync(writer, props);
                    widths = new ExcelWidthCollection(_configuration.MinWidth, _configuration.MaxWidth, props);
                }
                else
                {
                    await WriteColumnsWidthsAsync(writer, ExcelColumnWidth.FromProps(props));
                }

                await writer.WriteAsync(WorksheetXml.StartSheetData);
                int fieldCount = reader.FieldCount;
                if (_printHeader)
                {
                    await PrintHeaderAsync(writer, props);
                    yIndex++;
                }

                while (reader.Read())
                {
                    await writer.WriteAsync(WorksheetXml.StartRow(yIndex));
                    var xIndex = 1;
                    for (int i = 0; i < fieldCount; i++)
                    {
                        object cellValue;

                        if (_configuration.DynamicColumnFirst)
                        {
                            var columnIndex = reader.GetOrdinal(props[i].Key.ToString());
                            cellValue = reader.GetValue(columnIndex);
                        }
                        else
                        {
                            cellValue = reader.GetValue(i);
                        }

                        await WriteCellAsync(writer, yIndex, xIndex, cellValue, props[i], widths);
                        xIndex++;
                    }
                    await writer.WriteAsync(WorksheetXml.EndRow);
                    yIndex++;
                }

                // Subtract 1 because cell indexing starts with 1
                maxRowIndex = yIndex - 1;
            }

            await writer.WriteAsync(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            await writer.WriteAndFlushAsync(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode)
            {
                await WriteDimensionAsync(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }
            if (_configuration.EnableAutoWidth)
            {
                await OverWriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths.Columns);
            }
        }

        private async Task GenerateSheetByEnumerableAsync(MiniExcelAsyncStreamWriter writer, IEnumerable values)
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
                    await WriteEmptySheetAsync(writer);
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

            await writer.WriteAsync(WorksheetXml.StartWorksheetWithRelationship);

            long dimensionPlaceholderPostition = 0;

            // We can write the dimensions directly if the row count is known
            if (_configuration.FastMode && rowCount == null)
            {
                dimensionPlaceholderPostition = await WriteDimensionPlaceholderAsync(writer);
            }
            else
            {
                maxRowIndex = rowCount.Value + (_printHeader && rowCount > 0 ? 1 : 0);
                await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, maxColumnIndex)));
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
            var yIndex = 1;
            var xIndex = 1;
            if (_printHeader)
            {
                await PrintHeaderAsync(writer, props);
                yIndex++;
            }

            if (!empty)
            {
                // body
                switch (mode)
                {
                    case "IDictionary<string, object>": //Dapper Row
                        maxRowIndex = await GenerateSheetByColumnInfoAsync<IDictionary<string, object>>(writer, enumerator, props, widths, xIndex, yIndex);
                        break;
                    case "IDictionary":
                        maxRowIndex = await GenerateSheetByColumnInfoAsync<IDictionary>(writer, enumerator, props, widths, xIndex, yIndex);
                        break;
                    case "Properties":
                        maxRowIndex = await GenerateSheetByColumnInfoAsync<object>(writer, enumerator, props, widths, xIndex, yIndex);
                        break;
                    default:
                        throw new NotImplementedException($"Type {values.GetType().FullName} is not implemented. Please open an issue.");
                }
            }

            await writer.WriteAsync(WorksheetXml.EndSheetData);
            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }

            await writer.WriteAsync(WorksheetXml.Drawing(currentSheetIndex));
            await writer.WriteAsync(WorksheetXml.EndWorksheet);

            // The dimension has already been written if row count is defined
            if (_configuration.FastMode && rowCount == null)
            {
                await WriteDimensionAsync(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }
            if (_configuration.EnableAutoWidth)
            {
                await OverWriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths.Columns);
            }
        }

        private async Task GenerateSheetByDataTableAsync(MiniExcelAsyncStreamWriter writer, DataTable value)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            await writer.WriteAsync(WorksheetXml.StartWorksheet);
            var yIndex = xy.Item2;

            // dimension
            var maxRowIndex = value.Rows.Count + (_printHeader && value.Rows.Count > 0 ? 1 : 0);
            var maxColumnIndex = value.Columns.Count;
            await writer.WriteAsync(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, maxColumnIndex)));

            var props = new List<ExcelColumnInfo>();
            for (var i = 0; i < value.Columns.Count; i++)
            {
                var columnName = value.Columns[i].Caption ?? value.Columns[i].ColumnName;
                var prop = GetColumnInfosFromDynamicConfiguration(columnName);
                props.Add(prop);
            }

            //sheet view
            await writer.WriteAsync(GetSheetViews());

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

            await writer.WriteAsync(WorksheetXml.StartSheetData);
            if (_printHeader)
            {
                await writer.WriteAsync(WorksheetXml.StartRow(yIndex));
                var xIndex = xy.Item1;
                foreach (var p in props)
                {
                    var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                    await WriteCellAsync(writer, r, columnName: p.ExcelColumnName);
                    xIndex++;
                }

                await writer.WriteAsync(WorksheetXml.EndRow);
                yIndex++;
            }

            for (int i = 0; i < value.Rows.Count; i++)
            {
                await writer.WriteAsync(WorksheetXml.StartRow(yIndex));
                var xIndex = xy.Item1;

                for (int j = 0; j < value.Columns.Count; j++)
                {
                    var cellValue = value.Rows[i][j];
                    await WriteCellAsync(writer, yIndex, xIndex, cellValue, props[j], widths);
                    xIndex++;
                }
                await writer.WriteAsync(WorksheetXml.EndRow);
                yIndex++;
            }

            await writer.WriteAsync(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
            {
                await writer.WriteAsync(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }
            if (_configuration.EnableAutoWidth)
            {
                await OverWriteColumnWidthPlaceholdersAsync(writer, columnWidthsPlaceholderPosition, widths.Columns);
            }

            await writer.WriteAsync(WorksheetXml.EndWorksheet);
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

        private static async Task PrintHeaderAsync(MiniExcelAsyncStreamWriter writer, List<ExcelColumnInfo> props)
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

        private async Task<int> GenerateSheetByColumnInfoAsync<T>(MiniExcelAsyncStreamWriter writer, IEnumerator value, List<ExcelColumnInfo> props, ExcelWidthCollection widthCollection, int xIndex = 1, int yIndex = 1)
        {
            var isDic = typeof(T) == typeof(IDictionary);
            var isDapperRow = typeof(T) == typeof(IDictionary<string, object>);
            do
            {
                // The enumerator has already moved to the first item
                T v = (T)value.Current;

                await writer.WriteAsync(WorksheetXml.StartRow(yIndex));
                var cellIndex = xIndex;
                foreach (var p in props)
                {
                    if (p == null) //reason:https://github.com/shps951023/MiniExcel/issues/142
                    {
                        cellIndex++;
                        continue;
                    }

                    object cellValue = null;
                    if (isDic)
                    {
                        cellValue = ((IDictionary)v)[p.Key];
                        //WriteCell(writer, yIndex, cellIndex, cellValue, null); // why null because dictionary that needs to check type every time
                        //TODO: user can specefic type to optimize efficiency
                    }
                    else if (isDapperRow)
                    {
                        cellValue = ((IDictionary<string, object>)v)[p.Key.ToString()];
                    }
                    else
                    {
                        cellValue = p.Property.GetValue(v);
                    }

                    await WriteCellAsync(writer, yIndex, cellIndex, cellValue, p, widthCollection);

                    cellIndex++;
                }

                await writer.WriteAsync(WorksheetXml.EndRow);
                yIndex++;
            } while (value.MoveNext());

            return yIndex - 1;
        }

        private static async Task WriteCellAsync(MiniExcelAsyncStreamWriter writer, string cellReference, string columnName)
        {
            await writer.WriteAsync(WorksheetXml.Cell(cellReference, "str", "1", ExcelOpenXmlUtils.EncodeXML(columnName)));
        }

        private async Task WriteCellAsync(MiniExcelAsyncStreamWriter writer, int rowIndex, int cellIndex, object value, ExcelColumnInfo p, ExcelWidthCollection widthCollection)
        {
            var columnReference = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, rowIndex);
            var valueIsNull = value is null || value is DBNull;

            if (_configuration.EnableWriteNullValueCell && valueIsNull)
            {
                await writer.WriteAsync(WorksheetXml.EmptyCell(columnReference, "2"));
                return;
            }

            var tuple = GetCellValue(rowIndex, cellIndex, value, p, valueIsNull);

            var styleIndex = tuple.Item1;
            var dataType = tuple.Item2;
            var cellValue = tuple.Item3;
            var columnType = p.ExcelColumnType;

            /*Prefix and suffix blank space will lost after SaveAs #294*/
            var preserveSpace = cellValue != null && (cellValue.StartsWith(" ", StringComparison.Ordinal) ||
                                                      cellValue.EndsWith(" ", StringComparison.Ordinal));

            if (p.CustomFormatter != null)
            {
                try
                {
                    cellValue = p.CustomFormatter(cellValue);
                }
                catch (Exception e)
                {
                    //ignored
                }
            }
            
            await writer.WriteAsync(WorksheetXml.Cell(columnReference, dataType, styleIndex, cellValue, preserveSpace: preserveSpace, columnType: columnType));
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
            var styleXml = GetStylesXml(_configuration.DynamicColumns);

            await CreateZipEntryAsync(
                ExcelFileNames.Styles,
                ExcelContentTypes.Styles,
                styleXml,
                cancellationToken);
        }

        private async Task GenerateDrawinRelXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingRelationshipXml(sheetIndex);
                await CreateZipEntryAsync(
                    ExcelFileNames.DrawingRels(sheetIndex),
                    string.Empty,
                    ExcelXml.DefaultDrawingXmlRels.Replace("{{format}}", drawing),
                    cancellationToken);
            }
        }

        private async Task GenerateDrawingXmlAsync(CancellationToken cancellationToken)
        {
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                var drawing = GetDrawingXml(sheetIndex);
                await CreateZipEntryAsync(
                    ExcelFileNames.Drawing(sheetIndex),
                    ExcelContentTypes.Drawing,
                    ExcelXml.DefaultDrawing.Replace("{{format}}", drawing),
                    cancellationToken);
            }
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