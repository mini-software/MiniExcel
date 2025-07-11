using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.OpenXml.Models;
using MiniExcelLibs.Picture;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLibs;

public static partial class MiniExcel
{
    [CreateSyncVersion]
    public static async Task AddPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        using var stream = File.Open(path, FileMode.OpenOrCreate);
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task AddPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    }

    public static MiniExcelDataReader GetReader(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null)
    {
        var stream = FileHelper.OpenSharedRead(path);
        return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
    }

    public static MiniExcelDataReader GetReader(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null)
    {
        return new MiniExcelDataReader(stream, useHeaderRow, sheetName, excelType, startCell, configuration);
    }

    [CreateSyncVersion]
    public static async Task<int> InsertAsync(string path, object value, string? sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IMiniExcelConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

        if (!File.Exists(path))
        {
            var rowsWritten = await SaveAsAsync(path, value, printHeader, sheetName, excelType, configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        if (excelType == ExcelType.CSV)
        {
            using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan);
            return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
            return await InsertAsync(stream, value, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, printHeader, overwriteSheet, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope", Justification = "TODO: CsvWriter needs to be disposed")]
    public static async Task<int> InsertAsync(this Stream stream, object value, string? sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IMiniExcelConfiguration? configuration = null, bool printHeader = true, bool overwriteSheet = false, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);
        if (excelType == ExcelType.CSV)
        {
            var newValue = value is IEnumerable or IDataReader ? value : new[] { value };
            var provider = ExcelWriterFactory.GetProvider(stream, newValue, sheetName, excelType, configuration, false);
            return await provider.InsertAsync(overwriteSheet, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            configuration ??= new OpenXmlConfiguration { FastMode = true };
            return await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader).InsertAsync(overwriteSheet, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    public static async Task<int[]> SaveAsAsync(string path, object value, bool printHeader = true, string? sheetName = "Sheet1", ExcelType excelType = ExcelType.UNKNOWN, IMiniExcelConfiguration? configuration = null, bool overwriteFile = false, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
            throw new NotSupportedException("MiniExcel's SaveAs does not support the .xlsm format");

        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        return await SaveAsAsync(stream, value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    [SuppressMessage("Reliability", "CA2000:Dispose objects before losing scope", Justification = "TODO: CsvWriter needs to be disposed")]
    public static async Task<int[]> SaveAsAsync(this Stream stream, object value, bool printHeader = true, string? sheetName = "Sheet1", ExcelType excelType = ExcelType.XLSX, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        return await ExcelWriterFactory.GetProvider(stream, value, sheetName, excelType, configuration, printHeader)
            .SaveAsAsync(cancellationToken)
            .ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> QueryAsync<T>(string path, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, bool hasHeader = true, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var stream = FileHelper.OpenSharedRead(path);
        
        var query = QueryAsync<T>(stream, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, hasHeader, cancellationToken);
        await foreach (var item in query.ConfigureAwait(false))
            yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<T> QueryAsync<T>(this Stream stream, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, bool hasHeader = true, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, new()
    {
        using var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryAsync<T>(sheetName, startCell, hasHeader, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryAsync(useHeaderRow, sheetName, startCell, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(
                new ExpandoObject() as IDictionary<string, object?>, 
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
    /// <param name="cancellationToken"></param>
    /// <returns></returns>
    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryRangeAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, endCell, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", string endCell = "", IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startCell, endCell, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(
                new ExpandoObject() as IDictionary<string, object?>,
                (dict, p) => { dict.Add(p); return dict; });
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryRangeAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await foreach (var item in QueryRangeAsync(stream, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, configuration, cancellationToken).ConfigureAwait(false))
            yield return item;
    }

    [CreateSyncVersion]
    public static async IAsyncEnumerable<dynamic> QueryRangeAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, int startRowIndex = 1, int startColumnIndex = 1, int? endRowIndex = null, int? endColumnIndex = null, IMiniExcelConfiguration? configuration = null, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using var excelReader = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
        await foreach (var item in excelReader.QueryRangeAsync(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, cancellationToken).ConfigureAwait(false))
            yield return item.Aggregate(
                new ExpandoObject() as IDictionary<string, object?>,
                (dict, p) => { dict.Add(p); return dict; });
    }

    #endregion QueryRange

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(string path, string templatePath, object value, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(path);
        await SaveAsByTemplateAsync(stream, templatePath, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(path);
        await SaveAsByTemplateAsync(stream, templateBytes, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, string templatePath, object value, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        await ExcelTemplateFactory.GetProvider(stream, configuration)
            .SaveAsByTemplateAsync(templatePath, value, cancellationToken)
            .ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, byte[] templateBytes, object value, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        await ExcelTemplateFactory.GetProvider(stream, configuration)
            .SaveAsByTemplateAsync(templateBytes, value, cancellationToken)
            .ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    public static async Task SaveAsByTemplateAsync(this Stream stream, Stream templateStream, object value, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        await ExcelTemplateFactory.GetProvider(stream, configuration)
            .SaveAsByTemplateAsync(templateStream, value, cancellationToken)
            .ConfigureAwait(false);
    }

    #region MergeCells

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(string mergedFilePath, string path, ExcelType excelType = ExcelType.XLSX, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = File.Create(mergedFilePath);
        await MergeSameCellsAsync(stream, path, excelType, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(this Stream stream, string path, ExcelType excelType = ExcelType.XLSX, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        await ExcelTemplateFactory.GetProvider(stream, configuration, excelType)
            .MergeSameCellsAsync(path, cancellationToken)
            .ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task MergeSameCellsAsync(this Stream stream, byte[] filePath, ExcelType excelType = ExcelType.XLSX, IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        await ExcelTemplateFactory.GetProvider(stream, configuration, excelType)
            .MergeSameCellsAsync(filePath, cancellationToken)
            .ConfigureAwait(false);
    }

    #endregion

    /// <summary>
    /// QueryAsDataTable is not recommended, because it'll load all data into memory.
    /// </summary>
    [Obsolete("QueryAsDataTable is not recommended, because it'll load all data into memory.")]
    [CreateSyncVersion]
    public static async Task<DataTable> QueryAsDataTableAsync(string path, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await QueryAsDataTableAsync(stream, useHeaderRow, sheetName, excelType: ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<DataTable> QueryAsDataTableAsync(this Stream stream, bool useHeaderRow = true, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        /*Issue #279*/
        if (sheetName is null && excelType != ExcelType.CSV)
            sheetName = (await stream.GetSheetNamesAsync(configuration as OpenXmlConfiguration, cancellationToken).ConfigureAwait(false)).First();

        var dt = new DataTable(sheetName);
        var first = true;
        using var provider = await ExcelReaderFactory.GetProviderAsync(stream, ExcelTypeHelper.GetExcelType(stream, excelType), configuration, cancellationToken).ConfigureAwait(false);
        var rows = provider.QueryAsync(false, sheetName, startCell, cancellationToken);

        var columnDict = new Dictionary<string, string>();
        await foreach (var row in rows.ConfigureAwait(false))
        {
            if (first)
            {
                foreach (var entry in row)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var columnName = useHeaderRow ? entry.Value?.ToString() : entry.Key;
                    if (!string.IsNullOrWhiteSpace(columnName)) // avoid #298 : Column '' does not belong to table
                    {
                        var column = new DataColumn(columnName, typeof(object)) { Caption = columnName };
                        dt.Columns.Add(column);
                        columnDict.Add(entry.Key, columnName!);//same column name throw exception???
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

    [CreateSyncVersion]
    public static async Task<List<string>> GetSheetNamesAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetNamesAsync(stream, config, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<List<string>> GetSheetNamesAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        config ??= OpenXmlConfiguration.DefaultConfig;

        // todo: figure out why adding using statement breaks the tests
#pragma warning disable CA2000 // Dispose objects before losing scope
        var archive = new ExcelOpenXmlZip(stream);
#pragma warning restore CA2000 // Dispose objects before losing scope
        
        using var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        
        return rels?.Select(s => s.Name).ToList() ?? [];
    }

    [CreateSyncVersion]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(string path, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetInformationsAsync(stream, config, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<List<SheetInfo>> GetSheetInformationsAsync(this Stream stream, OpenXmlConfiguration? config = null, CancellationToken cancellationToken = default)
    {
        config ??= OpenXmlConfiguration.DefaultConfig;

        using var archive = new ExcelOpenXmlZip(stream);
        using var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, config, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await reader.GetWorkbookRelsAsync(archive.EntryCollection, cancellationToken).ConfigureAwait(false);
        
        return rels?.Select((s, i) => s.ToSheetInfo((uint)i)).ToList() ?? [];
    }

    [CreateSyncVersion]
    public static async Task<ICollection<string>> GetColumnsAsync(string path, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetColumnsAsync(stream, useHeaderRow, sheetName, excelType, startCell, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<ICollection<string>> GetColumnsAsync(this Stream stream, bool useHeaderRow = false, string? sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, string startCell = "A1", IMiniExcelConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
#pragma warning disable CA2007 // Consider calling ConfigureAwait on the awaited task
        await using var enumerator = QueryAsync(stream, useHeaderRow, sheetName, excelType, startCell, configuration, cancellationToken).GetAsyncEnumerator(cancellationToken);
#pragma warning restore CA2007 // Consider calling ConfigureAwait on the awaited task
        
        _ = enumerator.ConfigureAwait(false);
        if (!await enumerator.MoveNextAsync().ConfigureAwait(false))
            return [];
        
        return (enumerator.Current as IDictionary<string, object>)?.Keys ?? [];
    }

    [CreateSyncVersion]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(string path, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        return await GetSheetDimensionsAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<IList<ExcelRange>> GetSheetDimensionsAsync(this Stream stream, CancellationToken cancellationToken = default)
    {
        using var reader = await ExcelOpenXmlSheetReader.CreateAsync(stream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        return await reader.GetDimensionsAsync(cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(string csv, string xlsx, CancellationToken cancellationToken = default)
    {
        using var csvStream = FileHelper.OpenSharedRead(csv);
        using var xlsxStream = new FileStream(xlsx, FileMode.CreateNew);
        await ConvertCsvToXlsxAsync(csvStream, xlsxStream, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(Stream csv, Stream xlsx, CancellationToken cancellationToken = default)
    {
        var value = QueryAsync(csv, useHeaderRow: false, excelType: ExcelType.CSV, cancellationToken: cancellationToken).ConfigureAwait(false);
        await SaveAsAsync(xlsx, value, printHeader: false, excelType: ExcelType.XLSX, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(string xlsx, string csv, CancellationToken cancellationToken = default)
    {
        using var xlsxStream = FileHelper.OpenSharedRead(xlsx);
        using var csvStream = new FileStream(csv, FileMode.CreateNew);
        await ConvertXlsxToCsvAsync(xlsxStream, csvStream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(Stream xlsx, Stream csv, CancellationToken cancellationToken = default)
    {
        var value = QueryAsync(xlsx, useHeaderRow: false, excelType: ExcelType.XLSX, cancellationToken: cancellationToken).ConfigureAwait(false);
        await SaveAsAsync(csv, value, printHeader: false, excelType: ExcelType.CSV, cancellationToken: cancellationToken).ConfigureAwait(false);
    }
}