using MiniExcelLibs.OpenXml;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.Picture;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    public static partial class MiniExcel
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
        {
            using (var stream = File.Open(path, FileMode.OpenOrCreate))
                await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task AddPicture(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
        {
            await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
        }

        public static MiniExcelDataReader GetReader(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            var stream = FileHelper.OpenSharedRead(path);
            return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static MiniExcelDataReader GetReader(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null)
        {
            return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<int> InsertAsync(string path, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

            if (!File.Exists(path))
            {
                var rowsWritten = await SaveAsAsync(path, value, printHeader, sheetName, excelType, configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
                return rowsWritten.FirstOrDefault();
            }

            if (excelType == ExcelType.CSV)
            {
                using (var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan))
                    return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task <int> InsertAsync(this Stream stream, object value, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
        {
            stream.Seek(0, SeekOrigin.End);
            if (excelType == ExcelType.CSV)
            {
                var newValue = value is IEnumerable || value is IDataReader ? value : new[] { value };
                return await ExcelWriterFactory.GetProvider(stream, newValue, sheetName, excelType, configuration, false).InsertAsync(overwriteSheet, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                var configOrDefault = configuration ?? new OpenXmlConfiguration { FastMode = true };
                return await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configOrDefault, printHeader).InsertAsync(overwriteSheet, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<int[]> SaveAsAsync(string path, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null, bool overwriteFile = false, CancellationToken cancellationToken = default)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel's SaveAs does not support the .xlsm format");

            using (var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew))
                return await SaveAsAsync(stream, value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<int[]> SaveAsAsync(this Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            return await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader).SaveAsAsync(cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<T> QueryAsync<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, bool hasHeader = true, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                await foreach (var item in QueryAsync<T>(stream, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, hasHeader, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<T> QueryAsync<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, bool hasHeader = true, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
        {
            using (var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false))
                await foreach (var item in excelReader.QueryAsync<T>(sheetName, startCell, hasHeader, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                await foreach (var item in QueryAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false))
                await foreach (var item in excelReader.QueryAsync(useHeaderRow, sheetName, startCell, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item.Aggregate(new ExpandoObject() as IDictionary<string, object>,
                            (dict, p) => { dict.Add(p); return dict; });
        }

        #region QueryRange

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
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                await foreach (var item in QueryRangeAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, endCell, configuration, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false))
                await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startCell, endCell, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item.Aggregate(new ExpandoObject() as IDictionary<string, object>,
                            (dict, p) => { dict.Add(p); return dict; });
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                await foreach(var item in QueryRangeAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IConfiguration configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            using (var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false))
                await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).WithCancellation(cancellationToken).ConfigureAwait(false))
                    yield return item.Aggregate(new ExpandoObject() as IDictionary<string, object>,
                            (dict, p) => { dict.Add(p); return dict; });
        }

        #endregion QueryRange

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            using (var stream = File.Create(path))
                await SaveAsByTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            using (var stream = File.Create(path))
                await SaveAsByTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templatePath, value, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration).SaveAsByTemplateAsync(templateBytes, value, cancellationToken).ConfigureAwait(false);
        }

        #region MergeCells

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task MergeSameCellsAsync(string mergedFilePath, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            using (var stream = File.Create(mergedFilePath))
                await MergeSameCellsAsync(stream, path, excelType, configuration, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task MergeSameCellsAsync(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration, excelType).MergeSameCellsAsync(path, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task MergeSameCellsAsync(this Stream stream, byte[] filePath, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            await ExcelTemplateFactory.GetProvider(stream, configuration, excelType).MergeSameCellsAsync(filePath, cancellationToken).ConfigureAwait(false);
        }

        #endregion

        /// <summary>
        /// QueryAsDataTable is not recommended, because it'll load all data into memory.
        /// </summary>
        [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async static Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
            {
                return await QueryAsDataTableAsync(stream, useHeaderRow, sheetName, excelType: ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async static Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            if (sheetName == null && excelType != ExcelType.CSV) /*Issue #279*/
                sheetName = (await stream.GetSheetNamesAsync(configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false)).First();

            var dt = new DataTable(sheetName);
            var first = true;
            var provider = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
            var rows = provider.QueryAsync(false, sheetName, startCell);

            var columnDict = new Dictionary<string, string>();
            await foreach (IDictionary<string, object> row in rows.WithCancellation(cancellationToken).ConfigureAwait(false))
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

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<List<string>> GetSheetNamesAsync(string path, OpenXmlConfiguration config = null, CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return await GetSheetNamesAsync(stream, config, cancellationToken).ConfigureAwait(false);
        }

       [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<List<string>> GetSheetNamesAsync(this Stream stream, OpenXmlConfiguration config = null, CancellationToken cancellationToken = default)
        {
            config = config ?? OpenXmlConfiguration.DefaultConfig;

            var archive = new ExcelOpenXmlZip(stream);
            var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
            var rels = await reader.GetWorkbookRelsAsync(archive.entries, cancellationToken).ConfigureAwait(false);
            return rels.Select(s => s.Name).ToList();
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, OpenXmlConfiguration config = null, CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return await GetSheetInformationsAsync(stream, config, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<List<SheetInfo>> GetSheetInformationsAsync(this Stream stream, OpenXmlConfiguration config = null, CancellationToken cancellationToken = default)
        {
            config = config ?? OpenXmlConfiguration.DefaultConfig;

            var archive = new ExcelOpenXmlZip(stream);
            var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
            var rels = await reader.GetWorkbookRelsAsync(archive.entries, cancellationToken).ConfigureAwait(false);
            return rels.Select((s, i) => s.ToSheetInfo((uint)i)).ToList();
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<ICollection<string>> GetColumnsAsync(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return await GetColumnsAsync(stream, useHeaderRow, sheetName, excelType, startCell, configuration, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<ICollection<string>> GetColumnsAsync(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IConfiguration configuration = null, CancellationToken cancellationToken = default)
        {
            var enumerator = QueryAsync(stream, useHeaderRow, sheetName, excelType, startCell, configuration).GetAsyncEnumerator(cancellationToken);
            _ = enumerator.ConfigureAwait(false);
            if (!await enumerator.MoveNextAsync())
            {
                return null;
            }
            return (enumerator.Current as IDictionary<string, object>)?.Keys;
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                return await GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(this Stream stream, CancellationToken cancellationToken = default)
        {
            var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
            return await reader.GetDimensionsAsync(cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task ConvertCsvToXlsxAsync(string csv, string xlsx, CancellationToken cancellationToken = default)
        {
            using (var csvStream = FileHelper.OpenSharedRead(csv))
            using (var xlsxStream = new FileStream(xlsx, FileMode.CreateNew))
            {
                await ConvertCsvToXlsxAsync(csvStream, xlsxStream, cancellationToken: cancellationToken).ConfigureAwait(false);
            }
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task ConvertCsvToXlsxAsync(Stream csv, Stream xlsx, CancellationToken cancellationToken = default)
        {
            var value = QueryAsync(csv, useHeaderRow: false, excelType: ExcelType.CSV);
            await SaveAsAsync(xlsx, value, printHeader: false, excelType: ExcelType.XLSX, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task ConvertXlsxToCsvAsync(string xlsx, string csv, CancellationToken cancellationToken = default)
        {
            using (var xlsxStream = FileHelper.OpenSharedRead(xlsx))
            using (var csvStream = new FileStream(csv, FileMode.CreateNew))
                await ConvertXlsxToCsvAsync(xlsxStream, csvStream, cancellationToken).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public static async Task ConvertXlsxToCsvAsync(Stream xlsx, Stream csv, CancellationToken cancellationToken = default)
        {
            var value = QueryAsync(xlsx, useHeaderRow: false, excelType: ExcelType.XLSX);
            await SaveAsAsync(csv, value, printHeader: false, excelType: ExcelType.CSV, cancellationToken: cancellationToken).ConfigureAwait(false);
        }
    }
}