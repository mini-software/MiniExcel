namespace MiniExcelLibs
{
    using Utils;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.Threading;
    using System.Threading.Tasks;

    public static partial class MiniExcel
    {
        public static async Task SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool overwriteFile = false,CancellationToken cancellationToken = default(CancellationToken))
        {
            await Task.Run(() => SaveAs(path, value, printHeader, sheetName, excelType, configuration, overwriteFile),cancellationToken).ConfigureAwait(false);
        }

        public static async Task SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader).SaveAsAsync(cancellationToken);
        }
        
        public static async Task MergeSameCellsAsync(string mergedFilePath, string path, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await Task.Run(() => MergeSameCells(mergedFilePath, path, excelType, configuration) ,cancellationToken).ConfigureAwait(false);
        }

        public static async Task MergeSameCellsAsync(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, CancellationToken cancellationToken = default(CancellationToken))
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration, excelType).MergeSameCellsAsync(path, cancellationToken);
        }

        public static async Task<IEnumerable<dynamic>> QueryAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            return await Task.Run(() => Query(path, useHeaderRow, sheetName, excelType, startCell, configuration),cancellationToken);
        }

        public static async Task<IEnumerable<T>> QueryAsync<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken)) where T : class, new()
        {
            return await ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration).QueryAsync<T>(sheetName, startCell,cancellationToken).ConfigureAwait(false);
        }

        public static async Task<IEnumerable<T>> QueryAsync<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken)) where T : class, new()
        {
            return await Task.Run(() => Query<T>(path, sheetName, excelType, startCell, configuration),cancellationToken).ConfigureAwait(false);
        }

        public static async Task<IEnumerable<dynamic>> QueryAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            TaskCompletionSource<IEnumerable<dynamic>> tcs = new TaskCompletionSource<IEnumerable<dynamic>>();
            cancellationToken.Register(() => {
                tcs.TrySetCanceled();
            });

            await Task.Run(() =>
            {
                try
                {
                    tcs.TrySetResult(Query(stream, useHeaderRow, sheetName, excelType, startCell, configuration));
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }, cancellationToken);

            return await tcs.Task;
            
        }
        public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templatePath, value,cancellationToken).ConfigureAwait(false);
        }

        public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templateBytes, value,cancellationToken).ConfigureAwait(false);
        }

        public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await Task.Run(() => SaveAsByTemplate(path, templatePath, value, configuration),cancellationToken).ConfigureAwait(false);
        }

        public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            await Task.Run(() => SaveAsByTemplate(path, templateBytes, value, configuration),cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")] 
        public static async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            return await Task.Run(() => QueryAsDataTable(path, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration),cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        public static async Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null,CancellationToken cancellationToken = default(CancellationToken))
        {
            return await Task.Run(() => QueryAsDataTable(stream, useHeaderRow, sheetName, excelType, startCell, configuration),cancellationToken).ConfigureAwait(false);
        }
    }
}
