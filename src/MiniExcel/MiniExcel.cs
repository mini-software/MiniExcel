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
            ExcelWriterFactory.GetProvider(stream, excelType).SaveAs(value, printHeader);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query<T>(stream, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query<T>(sheetName);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName);
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in GetSheetNames(stream))
                    yield return item;
        }

        public static IEnumerable<string> GetSheetNames(this Stream stream)
        {
            var archive = new ExcelOpenXmlZip(stream);
            foreach (var item in ExcelOpenXmlSheetReader.GetWorkbookRels(archive.Entries))
                yield return item.Name;
        }
    }
}
