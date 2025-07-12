using MiniExcelLib.OpenXml.Picture;

namespace MiniExcelLib;

public sealed partial class OpenXmlExporter
{
    [CreateSyncVersion]
    public async Task AddExcelPictureAsync(string path, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        using var stream = File.Open(path, FileMode.OpenOrCreate);
        await MiniExcelPictureImplement.AddPictureAsync(stream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task AddExcelPictureAsync(Stream excelStream, CancellationToken cancellationToken = default, params MiniExcelPicture[] images)
    {
        await MiniExcelPictureImplement.AddPictureAsync(excelStream, cancellationToken, images).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> InsertExcelSheetAsync(string path, object value, string? sheetName = "Sheet1", bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's Insert does not support the .xlsm format");

        if (!File.Exists(path))
        {
            var rowsWritten = await ExportExcelAsync(path, value, printHeader, sheetName, configuration: configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
        return await InsertExcelSheetAsync(stream, value, sheetName, printHeader, overwriteSheet, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> InsertExcelSheetAsync(Stream stream, object value, string? sheetName = "Sheet1", 
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
    public async Task<int[]> ExportExcelAsync(string path, object value, bool printHeader = true, 
        string? sheetName = "Sheet1", bool overwriteFile = false, OpenXmlConfiguration? configuration = null, 
        CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's SaveAs does not support the .xlsm format");

        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        return await ExportExcelAsync(stream, value, printHeader, sheetName, configuration, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportExcelAsync(Stream stream, object value, bool printHeader = true, string? sheetName = "Sheet1", 
        OpenXmlConfiguration? configuration = null, CancellationToken cancellationToken = default)
    {
        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(cancellationToken).ConfigureAwait(false);
    }
}