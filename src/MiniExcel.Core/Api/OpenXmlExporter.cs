// ReSharper disable once CheckNamespace
namespace MiniExcelLib.Core;

public sealed partial class OpenXmlExporter
{
    internal OpenXmlExporter() { }
    
    
    [CreateSyncVersion]
    public async Task<int> InsertSheetAsync(string path, object value, string? sheetName = "Sheet1",
        bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null,
        IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's InsertSheet does not support the .xlsm format");

        if (!File.Exists(path))
        {
            var rowsWritten = await ExportAsync(path, value, printHeader, sheetName, configuration: configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
            return rowsWritten.FirstOrDefault();
        }

        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
        return await InsertSheetAsync(stream, value, sheetName, printHeader, overwriteSheet, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int> InsertSheetAsync(Stream stream, object value, string? sheetName = "Sheet1", 
        bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null, 
        IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);
        configuration ??= new OpenXmlConfiguration { FastMode = true };

        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.InsertAsync(overwriteSheet, cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportAsync(string path, object value, bool printHeader = true, 
        string? sheetName = "Sheet1", bool overwriteFile = false, OpenXmlConfiguration? configuration = null, 
        IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        if (Path.GetExtension(path).Equals(".xlsm", StringComparison.InvariantCultureIgnoreCase))
            throw new NotSupportedException("MiniExcel's Export does not support the .xlsm format");
        
        var filePath = path.EndsWith(".xlsx",  StringComparison.InvariantCultureIgnoreCase) ? path : $"{path}.xlsx" ;
        
        using var stream = overwriteFile ? File.Create(filePath) : new FileStream(filePath, FileMode.CreateNew);
        return await ExportAsync(stream, value, printHeader, sheetName, configuration, progress, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task<int[]> ExportAsync(Stream stream, object value, bool printHeader = true, string? sheetName = "Sheet1", 
        OpenXmlConfiguration? configuration = null, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.SaveAsAsync(progress, cancellationToken).ConfigureAwait(false);
    }
}