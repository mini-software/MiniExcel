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

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel SaveAs not support xlsm");
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, sheetName, GetExcelType(path, excelType), configuration);
        }

        /// <summary>
        /// Default SaveAs Xlsx file
        /// </summary>
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidDataException("Sheet name can not be empty or null");
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            ExcelWriterFactory.GetProvider(stream, excelType).SaveAs(value, sheetName, printHeader, configuration);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = Helpers.OpenSharedRead(path))
                foreach (var item in Query<T>(stream, sheetName, GetExcelType(path, excelType), startCell, configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query<T>(sheetName, startCell, configuration);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, GetExcelType(path, excelType), startCell, configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
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

        /// <summary>
        /// This method can avoid reading the template file every time
        /// </summary>
        /// <param name="stream">output stream</param>
        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
        {
            ExcelTemplateFactory.GetProvider(stream).SaveAsByTemplate(templateBytes, value);
        }

        /// <summary>
        /// This method is not recommended, because it'll load all data into memory.
        /// </summary>
        public static DataTable QueryAsDataTable(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = Helpers.OpenSharedRead(path))
                return QueryAsDataTable(stream, useHeaderRow, sheetName, GetExcelType(path, excelType), startCell, configuration);
        }

        /// <summary>
        /// This method is not recommended, because it'll load all data into memory.
        /// Note: if first row cell is null value, column type will be string
        /// </summary>
        public static DataTable QueryAsDataTable(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            if (sheetName == null)
                sheetName = stream.GetSheetNames().First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {
                    foreach (var key in row.Keys)
                    {
                        //TODO:base on cell t xml
                        // issue 223 : can't assign typeof object 
                        var type = row[key]?.GetType() ?? typeof(string);
                        if (type == null || type == typeof(object))
                            type = typeof(string);
                        dt.Columns.Add(key, type);
                    }

                    first = false;
                }

                var newRow = dt.NewRow();
                foreach (var key in row.Keys)
                {
                    if (row[key] == null)
                        newRow[key] = DBNull.Value;
                    else
                        newRow[key] = row[key];
                }
                    
                dt.Rows.Add(newRow);
            }
            return dt;
        }
    }
}
