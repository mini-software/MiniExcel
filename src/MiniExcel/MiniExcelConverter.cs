using MiniExcelLib.Core;
using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Csv;
using MiniExcelLib.OpenXml.Api;
using Zomp.SyncMethodGenerator;

namespace MiniExcelLib;

public static partial class MiniExcelConverter
{
    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(Stream csv, Stream xlsx, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        var value = MiniExcel.Importers
            .GetCsvImporter()
            .QueryAsync(csv, useHeaderRow: csvHasHeader, cancellationToken: cancellationToken);

        await MiniExcel.Exporters
            .GetOpenXmlExporter()
            .ExportAsync(xlsx, value, printHeader: csvHasHeader, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertCsvToXlsxAsync(string csvPath, string xlsx, bool csvHasHeader = false, CancellationToken cancellationToken = default)
    {
        using var csvStream = FileHelper.OpenSharedRead(csvPath);
        using var xlsxStream = new FileStream(xlsx, FileMode.CreateNew);

        await ConvertCsvToXlsxAsync(csvStream, xlsxStream, csvHasHeader, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(string xlsx, string csvPath, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        using var xlsxStream = FileHelper.OpenSharedRead(xlsx);
        using var csvStream = new FileStream(csvPath, FileMode.CreateNew);

        await ConvertXlsxToCsvAsync(xlsxStream, csvStream, xlsxHasHeader, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public static async Task ConvertXlsxToCsvAsync(Stream xlsx, Stream csv, bool xlsxHasHeader = true, CancellationToken cancellationToken = default)
    {
        var value = MiniExcel.Importers
            .GetOpenXmlImporter()
            .QueryAsync(xlsx, useHeaderRow: xlsxHasHeader, cancellationToken: cancellationToken)
            .ConfigureAwait(false);
        
        await MiniExcel.Exporters
            .GetCsvExporter()
            .ExportAsync(csv, value, printHeader: xlsxHasHeader, cancellationToken: cancellationToken).ConfigureAwait(false);
    }
}