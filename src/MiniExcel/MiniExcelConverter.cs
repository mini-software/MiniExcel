using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Csv;
using MiniExcelLib.OpenXml;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLib;

/// <summary>
/// Provides methods for converting between CSV and OpenXml (XLSX) documents.
/// </summary>
public static partial class MiniExcelConverter
{
    /// <summary>
    /// Converts a CSV file to an OpenXml (XLSX) file.
    /// </summary>
    /// <param name="csvPath">The file path to the CSV file to be converted.</param>
    /// <param name="xlsxPath">The file path where the resulting XLSX file will be created.</param>
    /// <param name="csvHasHeader">If true, the first row will be treated as headers and printed in the Excel file. Otherwise, all rows are treated as data. Default is false</param>
    /// <param name="cancellationToken"> A cancellation token to signal that the operation should be cancelled.</param>
    /// <returns>
    /// A task that completes when the OpenXml file has been successfully created.
    /// </returns>
    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(string csvPath, string xlsxPath, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        var csvStream = FileHelper.OpenSharedRead(csvPath);
        var xlsxStream = new FileStream(xlsxPath, FileMode.CreateNew);

#if SYNC_ONLY
        using var disposableCsvStream = csvStream;
        using var disposableXlsxStream = xlsxStream;
#else
        await using var disposableCsvStream = csvStream.ConfigureAwait(false);
        await using var disposableXlsxStream = xlsxStream.ConfigureAwait(false);
#endif

        await ConvertCsvToXlsxAsync(csvStream, xlsxStream, csvHasHeader, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts CSV data from a stream to OpenXml (XLSX) format in another stream asynchronously.
    /// </summary>
    /// <param name="csvStream">A readable stream containing CSV data.</param>
    /// <param name="xlsxStream">A writable stream where the Excel data will be written.</param>
    /// <param name="csvHasHeader">If true, the first row will be treated as headers and printed in the Excel file. Otherwise, all rows are treated as data. Default is false.</param>
    /// <param name="cancellationToken">A cancellation token to signal that the operation should be cancelled.</param>
    /// <returns>
    /// A task that completes when the OpenXml data has been successfully written to the output stream.
    /// </returns>
    /// <remarks>
    /// The streams will not be closed by this method; the caller is responsible for disposing them.
    /// </remarks>
    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(Stream csvStream, Stream xlsxStream, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        var value = MiniExcel.Importers.GetCsvImporter()
            .QueryAsync(csvStream, hasHeaderRow: csvHasHeader, leaveOpen: true, cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        await MiniExcel.Exporters.GetOpenXmlExporter()
            .ExportAsync(xlsxStream, value, printHeader: csvHasHeader, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts an OpenXml (XLSX) file to a CSV file.
    /// </summary>
    /// <param name="xlsxPath">The file path to the XLSX file to be converted.</param>
    /// <param name="csvPath">The file path where the resulting CSV file will be created.</param>
    /// <param name="xlsxHasHeader">If true, the first row will be treated as headers and printed in the CSV file. Otherwise, all rows are treated as data. Default is false.</param>
    /// <param name="cancellationToken"> A cancellation token to signal that the operation should be cancelled.</param>
    /// <returns>
    /// A task that completes when the CSV file has been successfully created.
    /// </returns>
    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(string xlsxPath, string csvPath, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        var xlsxStream = FileHelper.OpenSharedRead(xlsxPath);
        var csvStream = new FileStream(csvPath, FileMode.CreateNew);

#if SYNC_ONLY
        using var disposableXlsxStream = xlsxStream;
        using var disposableCsvStream = csvStream;
#else
        await using var disposableXlsxStream = xlsxStream.ConfigureAwait(false);
        await using var disposableCsvStream = csvStream.ConfigureAwait(false);
#endif

        await ConvertXlsxToCsvAsync(xlsxStream, csvStream, xlsxHasHeader, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts CSV data from a stream to OpenXml (XLSX) format in another stream asynchronously.
    /// </summary>
    /// <param name="xlsxStream">A readable stream containing the OpenXml data.</param>
    /// <param name="csvStream">A writable stream where the CSV data will be written.</param>
    /// <param name="xlsxHasHeader">If true, the first row will be treated as headers and printed in the Excel file. Otherwise, all rows are treated as data. Default is false.</param>
    /// <param name="cancellationToken">A cancellation token to signal that the operation should be cancelled.</param>
    /// <returns>
    /// A task that completes when the CSV data has been successfully written to the output stream.
    /// </returns>
    /// <remarks>
    /// The streams will not be closed by this method; the caller is responsible for disposing them.
    /// </remarks>
    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(Stream xlsxStream, Stream csvStream, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        var value = MiniExcel.Importers.GetOpenXmlImporter()
            .QueryAsync(xlsxStream, hasHeaderRow: xlsxHasHeader, leaveOpen: true, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        
        await MiniExcel.Exporters.GetCsvExporter()
            .ExportAsync(csvStream, value, printHeader: xlsxHasHeader, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }
}
