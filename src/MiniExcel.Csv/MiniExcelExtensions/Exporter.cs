using System.Collections;
using System.Data;
using MiniExcelLib.Core;
using MiniExcelLib.Core.Helpers;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLib.Csv.MiniExcelExtensions;

public static partial class Exporter
{
    #region Append / Export
    
    [CreateSyncVersion]
    public static async Task<int> AppendToCsvAsync(this MiniExcelExporter me, string path, object value, bool printHeader = true, 
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (!File.Exists(path))
        {
            var rowsWritten = await ExportCsvAsync(me, path, value, printHeader, false, configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        using var stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read, 4096, FileOptions.SequentialScan);
        return await me.AppendToCsvAsync(stream, value, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<int> AppendToCsvAsync(this MiniExcelExporter me, Stream stream, object value, CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);

        var newValue = value is IEnumerable or IDataReader ? value : new[] { value };

        using var writer = new CsvWriter(stream, newValue, false, configuration);
        return await writer.InsertAsync(false, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<int[]> ExportCsvAsync(this MiniExcelExporter me, string path, object value, bool printHeader = true, bool overwriteFile = false, 
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        return await ExportCsvAsync(me, stream, value, printHeader, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task<int[]> ExportCsvAsync(this MiniExcelExporter me, Stream stream, object value, bool printHeader = true, 
        CsvConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        using var writer = new CsvWriter(stream, value, printHeader, configuration);
        return await writer.SaveAsAsync(cancellationToken).ConfigureAwait(false);
    }

    #endregion
    
    #region Convert

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(this MiniExcelExporter me, Stream csv, Stream xlsx, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        var value = new MiniExcelImporter().QueryCsvAsync(csv, useHeaderRow: csvHasHeader, cancellationToken: cancellationToken);
        await me.ExportXlsxAsync(xlsx, value, printHeader: csvHasHeader, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(this MiniExcelExporter me, string csvPath, string xlsx, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        using var csvStream = FileHelper.OpenSharedRead(csvPath);
        using var xlsxStream = new FileStream(xlsx, FileMode.CreateNew);

        await me.ConvertCsvToXlsxAsync(csvStream, xlsxStream, csvHasHeader, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(this MiniExcelExporter me, string xlsx, string csvPath, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        using var xlsxStream = FileHelper.OpenSharedRead(xlsx);
        using var csvStream = new FileStream(csvPath, FileMode.CreateNew);

        await me.ConvertXlsxToCsvAsync(xlsxStream, csvStream, xlsxHasHeader, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(this MiniExcelExporter me, Stream xlsx, Stream csv, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        var value = new MiniExcelImporter().QueryXlsxAsync(xlsx, useHeaderRow: xlsxHasHeader, cancellationToken: cancellationToken).ConfigureAwait(false);
        await me.ExportCsvAsync(csv, value, printHeader: xlsxHasHeader, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    #endregion
}