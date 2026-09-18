using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.Editor;

/// <summary>
/// Represents a pipeline for editing Excel files with a fluent API.
/// </summary>
public sealed partial class OpenXmlEditingPipeline
{
    private readonly OpenXmlEditorInternals _internals;

    internal OpenXmlEditingPipeline(Stream stream, bool leaveOpen)
    {
        _internals = new OpenXmlEditorInternals(stream, leaveOpen);
    }

    /// <summary>
    /// Updates the style of a cell using a callback function that modifies the provided <see cref="OpenXmlCellStyle"/> object.
    /// </summary>
    /// <param name="cellReference">The cell reference in standard Excel format (e.g., "A1", "B5").</param>
    /// <param name="updateCellCallback">A callback function that receives an <see cref="OpenXmlCellStyle"/> object and applies the desired style changes.</param>
    /// <param name="sheetName">The name of the worksheet to update. If null or not specified, the first sheet in the workbook is used.</param>
    /// <returns>
    /// Returns this <see cref="OpenXmlEditingPipeline"/> instance to enable method chaining.
    /// </returns>
    /// <remarks>
    /// Modifications are queued in the pipeline and not written to the file until SaveChanges or SaveChangesAsync is called.
    /// </remarks>
    public OpenXmlEditingPipeline UpdateCellStyle(string cellReference, Action<OpenXmlCellStyle> updateCellCallback, string? sheetName = null)
    {
        if (updateCellCallback is null)
            throw new ArgumentNullException(nameof(updateCellCallback));

        var style = new OpenXmlCellStyle();
        updateCellCallback(style);

        return UpdateCellStyle(cellReference, style, sheetName);
    }

    /// <summary>
    /// Updates the style of a cell using a pre-configured <see cref="OpenXmlCellStyle"/> object.
    /// </summary>
    /// <param name="cellReference">The cell reference in standard Excel format (e.g., "A1", "B5").</param>
    /// <param name="cellStyle">The <see cref="OpenXmlCellStyle"/> object containing the style properties to apply to the cell.</param>
    /// <param name="sheetName"> The name of the worksheet to update. If null or not specified, the first sheet in the workbook is used.</param>
    /// <returns>
    /// Returns this <see cref="OpenXmlEditingPipeline"/> instance to enable method chaining.
    /// </returns>
    /// <remarks>
    /// Modifications are queued in the pipeline and not written to the file until SaveChanges or SaveChangesAsync is called.
    /// </remarks>
    public OpenXmlEditingPipeline UpdateCellStyle(string cellReference, OpenXmlCellStyle cellStyle, string? sheetName = null)
    {
        _internals.UpdateCellStyle(cellReference, cellStyle, sheetName);
        return this;
    }

    /// <summary>
    /// Applies all queued modifications to the Excel document and saves it to the original stream or file.
    /// The pipeline cannot be reused afterwards.
    /// </summary>
    /// <param name="cancellationToken">The token to monitor for cancellation requests.</param>
    /// <remarks>
    /// If the pipeline was created from a stream with <c>leaveOpen: true</c>, the caller is responsible for its disposal.
    /// </remarks>
    [CreateSyncVersion]
    public async Task SaveChangesAsync(CancellationToken cancellationToken = default)
    {
        await _internals.SaveAsync(cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Applies all queued modifications to the Excel document and saves it to the provided path.
    /// The pipeline cannot be reused afterwards.
    /// </summary>
    /// <param name="outputPath">The path to save the modified Excel document to.</param>
    /// <param name="cancellationToken">The token to monitor for cancellation requests.</param>
    /// <remarks>
    /// If the pipeline was created from a stream with <c>leaveOpen: true</c>, the caller is responsible for its disposal.
    /// </remarks>
    [CreateSyncVersion]
    public async Task SaveChangesAsync(string outputPath, CancellationToken cancellationToken = default)
    {
        var stream = File.OpenWrite(outputPath);
        await using var disposableStream = stream.ConfigureAwait(false);

        await SaveChangesAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Applies all queued modifications to the Excel document and saves it to the provided stream.
    /// The pipeline cannot be reused afterwards.
    /// </summary>
    /// <param name="outputStream">The stream to save the modified Excel document to.</param>
    /// <param name="cancellationToken">The token to monitor for cancellation requests.</param>
    /// <remarks>
    /// If the pipeline was created from a stream with <c>leaveOpen: true</c>, the caller is responsible for its disposal.
    /// The caller is always responsible for disposing the output stream.
    /// </remarks>
    [CreateSyncVersion]
    public async Task SaveChangesAsync(Stream outputStream, CancellationToken cancellationToken = default)
    {
        await _internals.SaveAsync(outputStream, cancellationToken).ConfigureAwait(false);
    }
}
