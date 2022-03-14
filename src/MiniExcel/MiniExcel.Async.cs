namespace MiniExcelLibs
{
    using MiniExcelLibs.Csv;
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
        public static Task SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return Task.Run(() => SaveAs(path, value, printHeader, sheetName, excelType , configuration));
        }

        public static Task SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            return ExcelWriterFactory.GetProvider(stream,value, sheetName, excelType, configuration, printHeader).SaveAsAsync();
        }

        public static Task<IEnumerable<dynamic>> QueryAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return Task.Run(() => Query(path, useHeaderRow, sheetName, excelType, startCell, configuration));
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType),configuration).QueryAsync<T>(sheetName, startCell);
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return Task.Run(() => Query<T>(path, sheetName, excelType, startCell, configuration));
        }

        public static async Task<IEnumerable<dynamic>> QueryAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return await ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration).QueryAsync(useHeaderRow, sheetName, startCell);
        }
        public static Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IConfiguration configuration = null)
        {
            return ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templatePath, value);
        }

        public static Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IConfiguration configuration = null)
        {
            return ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templateBytes, value);
        }

        public static Task SaveAsByTemplateAsync(string path, string templatePath, object value, IConfiguration configuration = null)
        {
            return Task.Run(() => SaveAsByTemplate(path, templatePath, value, configuration));
        }

        public static Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IConfiguration configuration = null)
        {
            return Task.Run(() => SaveAsByTemplate(path, templateBytes, value, configuration));
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        public static Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return Task.Run(() => QueryAsDataTable(path, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration));
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        public static Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return Task.Run(() => QueryAsDataTable(stream, useHeaderRow, sheetName, excelType, startCell, configuration));
        }
    }
}
