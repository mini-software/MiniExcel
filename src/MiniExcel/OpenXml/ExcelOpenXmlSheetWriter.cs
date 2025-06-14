using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.OpenXml.Styles;
using MiniExcelLibs.Utils;
using MiniExcelLibs.WriteAdapter;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly MiniExcelZipArchive _archive;
        private static readonly UTF8Encoding _utf8WithBom = new UTF8Encoding(true);
        private readonly OpenXmlConfiguration _configuration;
        private readonly Stream _stream;
        private readonly bool _printHeader;
        private readonly object _value;
        private readonly string _defaultSheetName;
        private readonly List<SheetDto> _sheets = new List<SheetDto>();
        private readonly List<FileDto> _files = new List<FileDto>();
        private int _currentSheetIndex = 0;

        public ExcelOpenXmlSheetWriter(Stream stream, object value, string sheetName, IConfiguration configuration, bool printHeader)
        {
            _stream = stream;
            // Why ZipArchiveMode.Update not ZipArchiveMode.Create?
            // R : Mode create - ZipArchiveEntry does not support seeking.'
            _configuration = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.DefaultConfig;
            if (_configuration.EnableAutoWidth && !_configuration.FastMode)
                throw new InvalidOperationException("Auto width requires fast mode to be enabled");

            var archiveMode = _configuration.FastMode ? ZipArchiveMode.Update : ZipArchiveMode.Create;
            _archive = new MiniExcelZipArchive(_stream, archiveMode, true, _utf8WithBom);

            _value = value;
            _printHeader = printHeader;
            _defaultSheetName = sheetName;
        }

        internal void GenerateDefaultOpenXml()
        {
            CreateZipEntry(ExcelFileNames.Rels, ExcelContentTypes.Relationships, ExcelXml.DefaultRels);
            CreateZipEntry(ExcelFileNames.SharedStrings, ExcelContentTypes.SharedStrings, ExcelXml.DefaultSharedString);
            GenerateStylesXml();
        }

        private int CreateSheetXml(object values, string sheetPath)
        {
            var entry = _archive.CreateEntry(sheetPath, CompressionLevel.Fastest);
            var rowsWritten = 0;
            
            using (var zipStream = entry.Open())
            using (var writer = new MiniExcelStreamWriter(zipStream, _utf8WithBom, _configuration.BufferSize))
            {
                if (values == null)
                {
                    WriteEmptySheet(writer);
                }
                else
                {
                    rowsWritten = WriteValues(writer, values);
                }
            }
            _zipDictionary.Add(sheetPath, new ZipPackageInfo(entry, ExcelContentTypes.Worksheet));
            return rowsWritten;
        }

        private static void WriteEmptySheet(MiniExcelStreamWriter writer)
        {
            writer.Write(ExcelXml.EmptySheetXml);
        }

        private static long WriteDimensionPlaceholder(MiniExcelStreamWriter writer)
        {
            var dimensionPlaceholderPostition = writer.WriteAndFlush(WorksheetXml.StartDimension);
            writer.Write(WorksheetXml.DimensionPlaceholder); // end of code will be replaced

            return dimensionPlaceholderPostition;
        }

        private static void WriteDimension(MiniExcelStreamWriter writer, int maxRowIndex, int maxColumnIndex, long placeholderPosition)
        {
            // Flush and save position so that we can get back again.
            var position = writer.Flush();

            writer.SetPosition(placeholderPosition);
            writer.WriteAndFlush($@"{GetDimensionRef(maxRowIndex, maxColumnIndex)}""");

            writer.SetPosition(position);
        }

        private int WriteValues(MiniExcelStreamWriter writer, object values)
        {
            var writeAdapter = MiniExcelWriteAdapterFactory.GetWriteAdapter(values, _configuration);

            var isKnownCount = writeAdapter.TryGetKnownCount(out var count);
            var props = writeAdapter.GetColumns();
            if (props == null)
            {
                WriteEmptySheet(writer);
                return 0;
            }
            
            int maxRowIndex;
            var maxColumnIndex = props.Count(x => x != null && !x.ExcelIgnore);

            writer.Write(WorksheetXml.StartWorksheetWithRelationship);

            long dimensionPlaceholderPostition = 0;

            // We can write the dimensions directly if the row count is known
            if (isKnownCount)
            {
                maxRowIndex = _printHeader ? count + 1 : count;
                writer.Write(WorksheetXml.Dimension(GetDimensionRef(maxRowIndex, maxColumnIndex)));
            }
            else if (_configuration.FastMode)
            {
                dimensionPlaceholderPostition = WriteDimensionPlaceholder(writer);
            }

            //sheet view
            writer.Write(GetSheetViews());

            //cols:width
            ExcelWidthCollection widths = null;
            long columnWidthsPlaceholderPosition = 0;
            if (_configuration.EnableAutoWidth)
            {
                columnWidthsPlaceholderPosition = WriteColumnWidthPlaceholders(writer, maxColumnIndex);
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
            }
            maxRowIndex = currentRowIndex;

            writer.Write(WorksheetXml.EndSheetData);

            if (_configuration.AutoFilter)
                writer.Write(WorksheetXml.Autofilter(GetDimensionRef(maxRowIndex, maxColumnIndex)));

            writer.Write(WorksheetXml.Drawing(_currentSheetIndex));
            writer.Write(WorksheetXml.EndWorksheet);

            if (_configuration.FastMode && dimensionPlaceholderPostition != 0)
            {
                WriteDimension(writer, maxRowIndex, maxColumnIndex, dimensionPlaceholderPostition);
            }
            if (_configuration.EnableAutoWidth)
            {
                OverwriteColumnWidthPlaceholders(writer, columnWidthsPlaceholderPosition, widths?.Columns);
            }

            if (_printHeader)
                maxRowIndex--;

            return maxRowIndex;
        }

        private static long WriteColumnWidthPlaceholders(MiniExcelStreamWriter writer, int count)
        {
            var placeholderPosition = writer.Flush();
            writer.WriteWhitespace(WorksheetXml.GetColumnPlaceholderLength(count));
            return placeholderPosition;
        }

        private static void OverwriteColumnWidthPlaceholders(MiniExcelStreamWriter writer, long placeholderPosition, IEnumerable<ExcelColumnWidth> columnWidths)
        {
            var position = writer.Flush();

            writer.SetPosition(placeholderPosition);
            WriteColumnsWidths(writer, columnWidths);

            writer.Flush();
            writer.SetPosition(position);
        }

        private static void WriteColumnsWidths(MiniExcelStreamWriter writer, IEnumerable<ExcelColumnWidth> columnWidths)
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

            if (hasWrittenStart)
            {
                writer.Write(WorksheetXml.EndCols);
            }
        }

        private void PrintHeader(MiniExcelStreamWriter writer, List<ExcelColumnInfo> props)
        {
            const int yIndex = 1;
            writer.Write(WorksheetXml.StartRow(yIndex));

            var xIndex = 1;
            foreach (var p in props)
            {
                //reason : https://github.com/mini-software/MiniExcel/issues/142
                if (p != null)
                {
                    if (p.ExcelIgnore)
                        continue;
                            
                    var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                    WriteCell(writer, r, columnName: p.ExcelColumnName);
                }
                xIndex++;
            }

            writer.Write(WorksheetXml.EndRow);
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
