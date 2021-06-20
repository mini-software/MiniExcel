namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Utils;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Threading.Tasks;

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel SaveAs not support xlsm");
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration);
        }

        public static Task SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return Task.Run(() => SaveAs(path, value, printHeader, sheetName, excelType , configuration));
        }

        public static void SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            GetWriterProvider(stream, sheetName, excelType).SaveAs(value, sheetName, printHeader, configuration);
        }

        public static Task SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            return GetWriterProvider(stream, sheetName, excelType).SaveAsAsync(value, sheetName, printHeader, configuration);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = Helpers.OpenSharedRead(path))
                foreach (var item in Query<T>(stream, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query<T>(sheetName, startCell, configuration);
        }

        public static Task<IEnumerable<dynamic>> QueryAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return Task.Run(() => Query(path, useHeaderRow, sheetName, excelType, startCell, configuration));
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).QueryAsync<T>(sheetName, startCell, configuration);
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return Task.Run(() => Query<T>(path, sheetName, excelType, startCell, configuration));
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
        }

        public static Task<IEnumerable<IDictionary<string, object>>> QueryAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return GetReaderProvider(stream, excelType).QueryAsync(useHeaderRow, sheetName, startCell, configuration);
        }

        public static List<string> GetSheetNames(string path)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                return GetSheetNames(stream);
        }

        public static List<string> GetSheetNames(this Stream stream)
        {
            var archive = new ExcelOpenXmlZip(stream);
            return ExcelOpenXmlSheetReader.GetWorkbookRels(archive.Entries).Select(s => s.Name).ToList();
        }

        public static ICollection<string> GetColumns(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                return GetColumns(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static ICollection<string> GetColumns(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return (Query(stream, useHeaderRow, sheetName, excelType, startCell, configuration).FirstOrDefault() as IDictionary<string, object>)?.Keys;
        }

        public static void SaveAsByTemplate(string path, string templatePath, object value)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templatePath, value);
        }

        public static void SaveAsByTemplate(string path, byte[] templateBytes, object value)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templateBytes, value);
        }

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
        {
            ExcelTemplateFactory.GetProvider(stream).SaveAsByTemplate(templatePath, value);
        }

        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
        {
            ExcelTemplateFactory.GetProvider(stream).SaveAsByTemplate(templateBytes, value);
        }

        public static Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value)
        {
            return ExcelTemplateFactory.GetProvider(stream).SaveAsByTemplateAsync(templatePath, value);
        }

        public static Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value)
        {
            return ExcelTemplateFactory.GetProvider(stream).SaveAsByTemplateAsync(templateBytes, value);
        }

        public static Task SaveAsByTemplateAsync(string path, string templatePath, object value)
        {
            return Task.Run(() => SaveAsByTemplate(path, templatePath, value));
        }

        public static Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value)
        {
            return Task.Run(() => SaveAsByTemplate(path, templateBytes, value));
        }


        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        public static DataTable QueryAsDataTable(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                return QueryAsDataTable(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration);
        }

        public static Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return Task.Run(() => QueryAsDataTable(path, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration));
        }
        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        public static DataTable QueryAsDataTable(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return ExcelOpenXmlSheetReader.QueryAsDataTableImpl(stream, useHeaderRow, ref sheetName, excelType, startCell, configuration);
        }

        private static IExcelWriterAsync GetWriterProvider(Stream stream, string sheetName, ExcelType excelType)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidDataException("Sheet name can not be empty or null");
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");

            return ExcelWriterFactory.GetProvider(stream, excelType);
        }

        private static IExcelReaderAsync GetReaderProvider(Stream stream, ExcelType excelType)
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType));
        }
    }
}
