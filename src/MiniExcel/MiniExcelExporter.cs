using MiniExcelLib.OpenXml.Picture;
using OpenXmlWriter = MiniExcelLib.OpenXml.OpenXmlWriter;

namespace MiniExcelLib;

public sealed partial class MiniExcelExporter
{
    [CreateSyncVersion]
    public async Task AddPictureXlsxAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        using var stream = File.Open(path, FileMode.OpenOrCreate);
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task AddPictureXlsxAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> InsertXlsxSheetAsync(string path, object value, string? sheetName = "Sheet1", bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

        if (!File.Exists(path))
        {
            var rowsWritten = await ExportXlsxAsync(path, value, printHeader, sheetName, configuration: configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
        return await InsertXlsxSheetAsync(stream, value, sheetName, printHeader, overwriteSheet, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> InsertXlsxSheetAsync(Stream stream, object value, string? sheetName = "Sheet1", 
        bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);
        configuration ??= new OpenXmlConfiguration { FastMode = true };

        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.InsertAsync(overwriteSheet, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportXlsxAsync(string path, object value, bool printHeader = true, 
        string? sheetName = "Sheet1", bool overwriteFile = false, OpenXmlConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's SaveAs does not support the .xlsm format");

        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        return await ExportXlsxAsync(stream, value, printHeader, sheetName, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportXlsxAsync(Stream stream, object value, bool printHeader = true, string? sheetName = "Sheet1", 
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(cancellationToken).ConfigureAwait(false);
    }
}