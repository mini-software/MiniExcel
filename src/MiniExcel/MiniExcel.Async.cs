namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Threading.Tasks;
    using Utils;

    public static partial class MiniExcel
    {
        public static async Task<int> InsertAsync(string path, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

            if (!File.Exists(path))
            {
                var rowsWritten = await SaveAsAsync(path, value, printHeader, sheetName, excelType, configuration, ct: cancellationToken);
                return rowsWritten.FirstOrDefault();
            }

            if (excelType == ExcelType.CSV)
            {
                using (var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken);
            }
            else
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken);
            }
        }

        public static async Task<int>InsertAsync(this Stream stream, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            stream.Seek(0, SeekOrigin.End);
            // reuse code
            if (excelType == ExcelType.CSV)
            {
                var newValue = value is IEnumerable || value is IDataReader ? value : new[]{value}.AsEnumerable();
                return await ExcelWriterFactory.GetProvider(stream, newValue, sheetName, excelType, configuration, false).InsertAsync(overwriteSheet, cancellationToken);
            }
            else
            {
                var configOrDefault = configuration ?? new OpenXmlConfiguration { FastMode = true };
                return await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configOrDefault, printHeader).InsertAsync(overwriteSheet, cancellationToken);
            }
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        public static async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => QueryAsDataTable(path, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration), cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        public static async Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => QueryAsDataTable(stream, useHeaderRow, sheetName, excelType, startCell, configuration), cancellationToken).ConfigureAwait(false);
        }
    }
}
