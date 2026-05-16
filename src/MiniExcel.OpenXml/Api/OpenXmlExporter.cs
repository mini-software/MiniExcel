// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed partial class OpenXmlExporter
{
    internal OpenXmlExporter() { }
    
    /// <summary>
    /// Inserts a new worksheet into an existing OpenXml document.
    /// </summary>
    /// <param name="path">The path to the OpenXml document to modify.</param>
    /// <param name="value">The data object to insert into the new sheet. This can be an enumerable collection of a reference type, a <see cref="DataTable" /> or a <see cref="IDataReader"/>.</param>
    /// <param name="sheetName">The name to assign to the new worksheet.</param>
    /// <param name="printHeader">If <c>true</c>, includes the header row in the new sheet; otherwise, only data rows are written.</param>
    /// <param name="overwriteSheet">If <c>true</c>, overwrites any existing sheet with the same name; otherwise, an exception will be raised if the sheet already exists.</param>
    /// <param name="configuration">Optional configuration settings for the insert operation.</param>
    /// <param name="progress">Optional progress reporter to track insertion progress. The report value represents the number of cells written.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>The number of rows written to the new sheet.</returns>
    /// <remarks>
    /// FastMode is automatically enabled for this process and disabling it will result in <see cref="InvalidOperationException"/>.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<int> InsertSheetAsync(string path, object value, string sheetName = "Sheet1",
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

    /// <summary>
    /// Inserts a new worksheet into an existing OpenXml document.
    /// </summary>
    /// <param name="stream">The stream containing the OpenXml document to modify.</param>
    /// <param name="value">The data object to insert into the new sheet. This can be an enumerable collection of a reference type, a <c>IEnumeable&lt;IDictionary&lt;string, object&gt;&gt;</c>, a <see cref="DataTable" /> or a <see cref="IDataReader"/>.</param>
    /// <param name="sheetName">The name to assign to the new worksheet.</param>
    /// <param name="printHeader">If <c>true</c>, includes the header row in the new sheet; otherwise, only data rows are written.</param>
    /// <param name="overwriteSheet">If <c>true</c>, overwrites any existing sheet with the same name; otherwise, an exception will be raised if the sheet already exists.</param>
    /// <param name="configuration">Optional configuration settings for the insert operation.</param>
    /// <param name="progress">Optional progress reporter to track insertion progress. The report value represents the number of cells written.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>The number of rows written to the new sheet.</returns>
    /// <remarks>
    /// The stream position is reset to the end before writing.
    /// FastMode is automatically enabled for this process and disabling it will result in <see cref="InvalidOperationException"/>.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<int> InsertSheetAsync(Stream stream, object value, string sheetName = "Sheet1", 
        bool printHeader = true, bool overwriteSheet = false, OpenXmlConfiguration? configuration = null, 
        IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        stream.Seek(0, SeekOrigin.End);
        configuration ??= new OpenXmlConfiguration { FastMode = true };

        var writer = await OpenXmlWriter
            .CreateAsync(stream, value, sheetName, printHeader, configuration, cancellationToken)
            .ConfigureAwait(false);
        
        return await writer.InsertAsync(overwriteSheet, progress, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Exports data to a file as an OpenXml document.
    /// </summary>
    /// <param name="path">The path to write the OpenXml document to.</param>
    /// <param name="value">The data object to export. This can be an enumerable collection of a reference type, a <c>IEnumeable&lt;IDictionary&lt;string, object&gt;&gt;</c>, a <see cref="DataTable" /> or a <see cref="IDataReader"/>.</param>
    /// <param name="printHeader">If <c>true</c>, includes the header row in the output; otherwise, only data rows are written.</param>
    /// <param name="sheetName">The name to assign to the worksheet.</param>
    /// <param name="overwriteFile">If <c>true</c>, overwrites the file at the specified path, otherwise a <see cref="IOException"/> will be raised if the file already exists.</param>
    /// <param name="configuration">Optional configuration settings for the export operation.</param>
    /// <param name="progress">Optional progress reporter to track export progress. The report value represents the number of cells written.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>An array of integers representing the number of rows written for each exported sheet.</returns>
    [CreateSyncVersion]
    public async Task<int[]> ExportAsync(string path, object value, bool printHeader = true, 
        string sheetName = "Sheet1", bool overwriteFile = false, OpenXmlConfiguration? configuration = null, 
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

    /// <summary>
    /// Exports data to a stream as an OpenXml document.
    /// </summary>
    /// <param name="stream">The stream to write the OpenXml document.</param>
    /// <param name="value">The data object to export. This can be an enumerable collection of a reference type, a <c>IEnumeable&lt;IDictionary&lt;string, object&gt;&gt;</c>, a <see cref="DataTable" /> or a <see cref="IDataReader"/>.</param>
    /// <param name="printHeader">If <c>true</c>, includes the header row in the output; otherwise, only data rows are written.</param>
    /// <param name="sheetName">The name to assign to the worksheet.</param>
    /// <param name="configuration">Optional configuration settings for the export operation.</param>
    /// <param name="progress">Optional progress reporter to track export progress. The report value represents the number of cells written.</param>
    /// <param name="cancellationToken">A cancellation token to monitor for cancellation requests.</param>
    /// <returns>An array of integers representing the number of rows written for each exported sheet.</returns>
    /// <remarks>
    /// The stream position is not reset before writing.
    /// </remarks>
    [CreateSyncVersion]
    public async Task<int[]> ExportAsync(Stream stream, object value, bool printHeader = true, string sheetName = "Sheet1", 
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
