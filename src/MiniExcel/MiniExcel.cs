namespace MiniExcelLibs
{
    using OpenXml;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.Dynamic;
    using System.IO;
    using System.Linq;
    using Utils;
    using Zip;

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

        public static int Insert(string path, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

            if (!File.Exists(path))
            {
                var rowsWritten = SaveAs(path, value, printHeader, sheetName, excelType, configuration);
                return rowsWritten.FirstOrDefault();
            }

            if (excelType == ExcelType.CSV)
            {
                using (var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return Insert(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet);
            }
            else
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return Insert(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet);
            }
        }

        public static int Insert(this Stream stream, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false)
        {
            stream.Seek(0, SeekOrigin.End);
            // reuse code
            if (excelType == ExcelType.CSV)
            {
                var newValue = value is IEnumerable || value is IDataReader ? value : new[]{value}.AsEnumerable();
                return ExcelWriterFactory.GetProvider(stream, newValue, sheetName, excelType, configuration, false).Insert(overwriteSheet);
            }
            else
            {
                var configOrDefault = configuration ?? new OpenXmlConfiguration { FastMode = true };
                return ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configOrDefault, printHeader).Insert(overwriteSheet);
            } 
        }

        public static int[] SaveAs(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool overwriteFile = false)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel's SaveAs does not support the .xlsm format");

            using (var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew))
                return SaveAs(stream, value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration);
        }

        public static int[] SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            if (sheetName.Length > 31 && excelType == ExcelType.XLSX)
                throw new ArgumentException("Sheet names must be less than 31 characters", nameof(sheetName));
            return ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader).SaveAs();
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
                    yield return item.Aggregate(new ExpandoObject() as IDictionary<string, object>,
                            (dict, p) => { dict.Add(p); return dict; });
        }

        #region range

        /// <summary>
        /// Extract the given range。 Only uppercase letters are effective。
        /// e.g.
        ///     MiniExcel.QueryRange(path, startCell: "A2", endCell: "C3")
        ///     A2 represents the second row of column A, C3 represents the third row of column C
        ///     If you don't want to restrict rows, just don't include numbers
        /// </summary>
        /// <param name="path"></param>
        /// <param name="useHeaderRow"></param>
        /// <param name="sheetName"></param>
        /// <param name="excelType"></param>
        /// <param name="startCell">top left corner</param>
        /// <param name="endCell">lower right corner</param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        public static IEnumerable<dynamic> QueryRange(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "a1", string endCell = "", IConfiguration configuration = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                foreach (var item in QueryRange(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell == "" ? "a1" : startCell, endCell, configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> QueryRange(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "a1", string endCell = "", IConfiguration configuration = null)
        {
            using (var excelReader = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration))
                foreach (var item in excelReader.QueryRange(useHeaderRow, sheetName, startCell == "" ? "a1" : startCell, endCell))
                    yield return item.Aggregate(new ExpandoObject() as IDictionary<string, object>,
                            (dict, p) => { dict.Add(p); return dict; });
        }

        #endregion range

        public static void SaveAsByTemplate(string path, string templatePath, object value, IConfiguration configuration = null)
        {
            using (var stream = File.Create(path))
                SaveAsByTemplate(stream, templatePath, value, configuration);
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

        #region MergeCells

        public static void MergeSameCells(string mergedFilePath, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            using (var stream = File.Create(mergedFilePath))
                MergeSameCells(stream, path, excelType, configuration);
        }

        public static void MergeSameCells(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            ExcelTemplateFactory.GetProvider(stream, configuration, excelType).MergeSameCells(path);
        }

        public static void MergeSameCells(this Stream stream, byte[] filePath, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            ExcelTemplateFactory.GetProvider(stream, configuration, excelType).MergeSameCells(filePath);
        }

        #endregion

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        public static DataTable QueryAsDataTable(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
            {
                return QueryAsDataTable(stream, useHeaderRow, sheetName, excelType: ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration);
            }
        }

        public static DataTable QueryAsDataTable(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            if (sheetName == null && excelType != ExcelType.CSV) /*Issue #279*/
                sheetName = stream.GetSheetNames(configuration as OpenXmlConfiguration).First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration).Query(false, sheetName, startCell);

            var columnDict = new Dictionary<string, string>();
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {
                    foreach (var entry in row)
                    {
                        var columnName = useHeaderRow ? entry.Value?.ToString() : entry.Key;
                        if (!string.IsNullOrWhiteSpace(columnName)) // avoid #298 : Column '' does not belong to table
                        {
                            var column = new DataColumn(columnName, typeof(object)) { Caption = columnName };
                            dt.Columns.Add(column);
                            columnDict.Add(entry.Key, columnName);//same column name throw exception???
                        }
                    }
                    dt.BeginLoadData();
                    first = false;
                    if (useHeaderRow)
                    {
                        continue;
                    }
                }

                var newRow = dt.NewRow();
                foreach (var entry in columnDict)
                {
                    newRow[entry.Value] = row[entry.Key]; //TODO: optimize not using string key
                }

                dt.Rows.Add(newRow);
            }

            dt.EndLoadData();
            return dt;
        }

        public static List<string> GetSheetNames(string path, OpenXmlConfiguration config = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return GetSheetNames(stream, config);
        }

        public static List<string> GetSheetNames(this Stream stream, OpenXmlConfiguration config = null)
        {
            config = config ?? OpenXmlConfiguration.DefaultConfig;

            var archive = new ExcelOpenXmlZip(stream);
            return new ExcelOpenXmlSheetReader(stream, config).GetWorkbookRels(archive.entries).Select(s => s.Name).ToList();
        }

        public static List<SheetInfo> GetSheetInformations(string path, OpenXmlConfiguration config = null)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return GetSheetInformations(stream, config);
        }

        public static List<SheetInfo> GetSheetInformations(this Stream stream, OpenXmlConfiguration config = null)
        {
            config = config ?? OpenXmlConfiguration.DefaultConfig;

            var archive = new ExcelOpenXmlZip(stream);
            return new ExcelOpenXmlSheetReader(stream, config).GetWorkbookRels(archive.entries).Select((s, i) => s.ToSheetInfo((uint)i)).ToList();
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