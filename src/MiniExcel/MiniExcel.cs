namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.IO;

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, sheetName, GetExcelType(path, excelType));
        }

        /// <summary>
        /// Default SaveAs Xlsx file
        /// </summary>
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = null, ExcelType excelType = ExcelType.XLSX)
        {
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            ExcelWriterFactory.GetProvider(stream,excelType).SaveAs(value, printHeader);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query<T>(stream, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream,GetExcelType(stream, excelType)).Query<T>(sheetName);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream,GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName);
        }

        public static GridReader QueryMultiple(string path, bool useHeaderRow = false)
        {
            // only xlsx support this
            return new GridReader(path);
        }
    }

    public class GridReader : IDisposable
    {
        private string _path;
        private List<SheetRecord> _sheetRecords;
        private int _sheetIndex = 0;

        public GridReader(string path)
        {
            _path = path;

            using (var stream = File.OpenRead(path))
            using (var archive = new ExcelOpenXmlZip(stream))
            {
                _sheetRecords = ExcelOpenXmlSheetReader.GetWorkbookRels(archive.Entries);
            }
        }

        public IEnumerable<string> GetSheetNames()
        {
            foreach (var item in _sheetRecords)
            {
                yield return item.Name;
            }
        }

        public void Dispose()
        {
            //if (archive != null)
            //{
            //    archive.Dispose();
            //    archive = null;
            //}
            GC.SuppressFinalize(this);
        }

        public IEnumerable<dynamic> Read(bool useHeaderRow = false)
        {
            var sheetRecord = _sheetRecords[_sheetIndex];//TODO
            _sheetIndex = _sheetIndex + 1;
            return MiniExcel.Query(_path,useHeaderRow, sheetRecord.Name);
        }
    }
}
