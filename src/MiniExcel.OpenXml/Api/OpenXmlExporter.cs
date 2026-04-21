namespace MiniExcelLib.OpenXml.Api;

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

#if NET8_0_OR_GREATER
        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, 4096, FileOptions.SequentialScan);
#endif
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

#if NET8_0_OR_GREATER
        var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = overwriteFile ? File.Create(path) : new FileStream(path, FileMode.CreateNew);
#endif
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

    /// <summary>
    /// Modify the properties of a worksheet in the specified document.
    /// </summary>
    /// <param name="path">The path to the OpenXml document.</param>
    /// <param name="sheetName">The name of the worksheet to modify.</param>
    /// <param name="newSheetName">The new name to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="newSheetIndex">The position in the workbook to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="newSheetState">The visibility state to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="cancellationToken">The token to monitor for cancellation requests</param>
    [CreateSyncVersion]
    public async Task AlterSheetAsync(string path, string sheetName, string? newSheetName = null, int? newSheetIndex = null, SheetState? newSheetState = null, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        await using var disposableStream = stream.ConfigureAwait(false); 
#else
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
#endif
        await AlterSheetAsync(stream, sheetName, newSheetName, newSheetIndex, newSheetState, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Modify the properties of a worksheet in the specified document.
    /// </summary>
    /// <param name="stream">The stream to the OpenXml document.</param>
    /// <param name="sheetName">The name of the worksheet to modify.</param>
    /// <param name="newSheetName">The new name to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="newSheetIndex">The position in the workbook to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="newSheetState">The visibility state to assign to the worksheet, or <c>null</c> to leave as is.</param>
    /// <param name="cancellationToken">The token to monitor for cancellation requests</param>
    [CreateSyncVersion]
    public async Task AlterSheetAsync(Stream stream, string sheetName, string? newSheetName = null, int? newSheetIndex = null, SheetState? newSheetState = null, CancellationToken cancellationToken = default)
    {
        var writer = await OpenXmlWriter
            .CreateAsync(stream, null, sheetName, false, new OpenXmlConfiguration { FastMode = true }, cancellationToken)
            .ConfigureAwait(false);

        await writer.AlterWorksheetAsync(sheetName, newSheetName, newSheetIndex, newSheetState, cancellationToken).ConfigureAwait(false);
    }
}
