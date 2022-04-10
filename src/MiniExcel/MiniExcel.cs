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
        public static MiniExcelDataReader GetReader(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            var stream = FileHelper.OpenSharedRead(path);
            return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static MiniExcelDataReader GetReader(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static void SaveAs(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null,bool overwriteFile = false)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel SaveAs not support xlsm");

            using (var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration);
        }

        public static void SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader).SaveAs();
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                foreach (var item in Query<T>(stream, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null) where T : class, new()
        {
            using (var excelReader = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration))
                foreach (var item in excelReader.Query<T>(sheetName, startCell))
                {
                    yield return item;
                }
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var excelReader = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration))
                foreach (var item in excelReader.Query(useHeaderRow, sheetName, startCell))
                    yield return item;
        }

        public static void SaveAsByTemplate(string path, string templatePath, object value, IConfiguration configuration = null)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templatePath, value,configuration);
        }

        public static void SaveAsByTemplate(string path, byte[] templateBytes, object value, IConfiguration configuration = null)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templateBytes, value, configuration);
        }

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value, IConfiguration configuration = null)
        {
            ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplate(templatePath, value);
        }

        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value, IConfiguration configuration = null)
        {
            ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplate(templateBytes, value);
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        public static DataTable QueryAsDataTable(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
            {
                return QueryAsDataTable(stream, useHeaderRow, sheetName, excelType:ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration);
            }
        }
        public static DataTable QueryAsDataTable(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            if (sheetName == null && excelType != ExcelType.CSV) /*Issue #279*/
                sheetName = stream.GetSheetNames().First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType),configuration).Query(useHeaderRow, sheetName, startCell);

            var keys = new List<string>();
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {
                    foreach (var key in row.Keys)
                    {
                        if (!string.IsNullOrEmpty(key)) // avoid #298 : Column '' does not belong to table
                        {
                            var column = new DataColumn(key, typeof(object)) { Caption = key };
                            dt.Columns.Add(column);
                            keys.Add(key);
                        }
                    }

                    dt.BeginLoadData();
                    first = false;
                }

                var newRow = dt.NewRow();
                foreach (var key in keys)
                {
                    newRow[key] = row[key]; //TODO: optimize not using string key
                }

                dt.Rows.Add(newRow);
            }

            dt.EndLoadData();
            return dt;
        }

        public static List<string> GetSheetNames(string path)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return GetSheetNames(stream);
        }

        public static List<string> GetSheetNames(this Stream stream)
        {
            var archive = new ExcelOpenXmlZip(stream);
            return new ExcelOpenXmlSheetReader(stream,null).GetWorkbookRels(archive.entries).Select(s => s.Name).ToList();
        }

        public static ICollection<string> GetColumns(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return GetColumns(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static ICollection<string> GetColumns(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return (Query(stream, useHeaderRow, sheetName, excelType, startCell, configuration).FirstOrDefault() as IDictionary<string, object>)?.Keys;
        }

        public static void ConvertCsvToXlsx(string csv, string xlsx)
        {
            using (var csvStream = FileHelper.OpenSharedRead(csv))
            using (var xlsxStream = new FileStream(xlsx, FileMode.CreateNew))
            {
                ConvertCsvToXlsx(csvStream, xlsxStream);
            }
        }

        public static void ConvertCsvToXlsx(Stream csv, Stream xlsx)
        {
            var value = Query(csv, useHeaderRow: false, excelType: ExcelType.CSV);
            SaveAs(xlsx, value, printHeader: false, excelType: ExcelType.XLSX);
        }

        public static void ConvertXlsxToCsv(string xlsx, string csv)
        {
            using (var xlsxStream = FileHelper.OpenSharedRead(xlsx))
            using (var csvStream = new FileStream(csv, FileMode.CreateNew))
                ConvertXlsxToCsv(xlsxStream, csvStream);
        }

        public static void ConvertXlsxToCsv(Stream xlsx, Stream csv)
        {
            var value = Query(xlsx, useHeaderRow: false, excelType: ExcelType.XLSX);
            SaveAs(csv, value, printHeader: false, excelType: ExcelType.CSV);
        }
    }
}
