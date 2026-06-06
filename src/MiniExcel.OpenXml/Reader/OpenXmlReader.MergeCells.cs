namespace MiniExcelLib.OpenXml.Reader;

internal partial class OpenXmlReader
{
    [CreateSyncVersion]
    internal static async Task<(bool Success, MergeCells? MergeCells)> TryGetMergeCellsAsync(ZipArchiveEntry sheetEntry, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var xmlSettings = XmlReaderHelper.GetXmlReaderSettings();
        var mergeCells = new MergeCells();

        var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await  using var disposableSheetStream = sheetStream.ConfigureAwait(false);

        using var reader = XmlReader.Create(sheetStream, xmlSettings);
        
        if (!reader.IsStartElement("worksheet", Ns))
            return (false, null);
        
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            if (!reader.IsStartElement("mergeCells", Ns))
                continue;

            if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                return (false, null);

            while (!reader.EOF)
            {
                if (reader.IsStartElement("mergeCell", Ns))
                {
                    var refAttr = reader.GetAttribute("ref");
                    if (refAttr?.Split(':') is not [var startCell, var endCell])
                        continue;

                    CellReferenceConverter.TryParseCellReference(startCell, out var x1, out var y1);
                    CellReferenceConverter.TryParseCellReference(endCell, out var x2, out var y2);

                    mergeCells.MergesValues.Add(startCell, null);

                    // foreach range
                    var isFirst = true;
                    for (int x = x1; x <= x2; x++)
                    {
                        for (int y = y1; y <= y2; y++)
                        {
                            if (!isFirst)
                                mergeCells.MergesMap.Add(CellReferenceConverter.GetCellFromCoordinates(x, y), startCell);
                            isFirst = false;
                        }
                    }

                    await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false);
                }
                else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }
        }

        return (true, mergeCells);
    }
}
