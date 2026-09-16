using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.Editor;

public sealed partial class OpenXmlEditingPipeline
{
    private readonly OpenXmlEditorInternals _internals;

    internal OpenXmlEditingPipeline(Stream stream, bool leaveOpen)
    {
        _internals = new OpenXmlEditorInternals(stream, leaveOpen);
    }

    /// <summary>Queues a partial style update for an existing cell.</summary>
    public OpenXmlEditingPipeline UpdateCellStyle(string cellReference, Action<OpenXmlCellStyle> updateCellFunc, string? sheetName = null)
    {
        if (updateCellFunc is null)
            throw new ArgumentNullException(nameof(updateCellFunc));

        var style = new OpenXmlCellStyle();
        updateCellFunc(style);

        return UpdateCellStyle(cellReference, style, sheetName);
    }

    /// <summary>Queues a partial style update for an existing cell.</summary>
    public OpenXmlEditingPipeline UpdateCellStyle(string cellReference, OpenXmlCellStyle cellStyle, string? sheetName = null)
    {
        _internals.UpdateCellStyle(cellReference, cellStyle, sheetName);
        return this;
    }

    [CreateSyncVersion]
    public async Task SaveChangesAsync(CancellationToken cancellationToken = default)
    {
        await _internals.SaveAsync(cancellationToken).ConfigureAwait(false);
    }
}
