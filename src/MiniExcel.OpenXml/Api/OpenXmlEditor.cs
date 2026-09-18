using MiniExcelLib.OpenXml.Editor;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed partial class OpenXmlEditor
{
    internal OpenXmlEditor() { }


    /// <summary>
    /// Creates a new editing pipeline for the provided Excel document.
    /// </summary>
    /// <param name="path">The file path to the Excel document to edit.</param>
    /// <returns>
    /// An <see cref="OpenXmlEditingPipeline"/> instance that can be used to apply modifications to the document.
    /// </returns>
    /// <remarks>
    /// This method opens the file for exclusive read-write access. The file is locked until 
    /// SaveChanges or SaveChangesAsync is called.
    /// </remarks>
    public OpenXmlEditingPipeline StartEditingPipeline(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path cannot be null or whitespace.", nameof(path));

        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        return new OpenXmlEditingPipeline(stream, leaveOpen: false);
    }

    /// <summary>
    /// Creates a new editing pipeline for the provided Excel document.
    /// </summary>
    /// <param name="stream">The stream containing the Excel file data.</param>
    /// <param name="leaveOpen">
    /// If true the stream remains open after changes are saved and must be disposed by the caller,
    /// if false it is automatically closed when changes are saved. Default is false.
    /// </param>
    /// <returns>
    /// An <see cref="OpenXmlEditingPipeline"/> instance that can be used to apply modifications to the file.
    /// </returns>
    /// <remarks>
    /// Even with parameter <c>leaveOpen: false</c>, the underlying stream will not be disposed until
    /// SaveChanges or SaveChangesAsync are called.
    /// </remarks>
    public OpenXmlEditingPipeline StartEditingPipeline(Stream stream, bool leaveOpen = false)
    {
        if (stream is null)
            throw new ArgumentNullException(nameof(stream));

        return new OpenXmlEditingPipeline(stream, leaveOpen);
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
    public async Task AlterSheetInfoAsync(string path, string sheetName, string? newSheetName = null, int? newSheetIndex = null, SheetState? newSheetState = null, CancellationToken cancellationToken = default)
    {
        var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        await using var disposableStream = stream.ConfigureAwait(false); 

        await AlterSheetInfoAsync(stream, sheetName, newSheetName, newSheetIndex, newSheetState, cancellationToken).ConfigureAwait(false);
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
    public async Task AlterSheetInfoAsync(Stream stream, string sheetName, string? newSheetName = null, int? newSheetIndex = null, SheetState? newSheetState = null, CancellationToken cancellationToken = default)
    {
        var internals = new OpenXmlEditorInternals(stream, true);
        await internals.AlterWorksheetAsync(sheetName, newSheetName, newSheetIndex, newSheetState, cancellationToken).ConfigureAwait(false);
    }
}
