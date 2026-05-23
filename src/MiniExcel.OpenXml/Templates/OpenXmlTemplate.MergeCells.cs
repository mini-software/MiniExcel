namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default)
    {
        var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var disposableStream = stream.ConfigureAwait(false);

        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default)
    {
        var stream = new MemoryStream(fileInBytes);
        await using var disposableStream = stream.ConfigureAwait(false);

        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task MergeSameCellsImplAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        await stream.CopyToAsync(_outputFileStream
#if NET
            , cancellationToken
#endif
        ).ConfigureAwait(false);

        using var reader = await OpenXmlReader.CreateAsync(_outputFileStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        var archive = await OpenXmlZip.CreateAsync(_outputFileStream, mode: ZipArchiveMode.Update, true, Encoding.UTF8, true, cancellationToken).ConfigureAwait(false);
        await using  var disposableArchive = archive.ConfigureAwait(false);

        //read sharedString
        var sharedStrings = reader.SharedStrings;

        //read all xlsx sheets
        var sheets = archive.ZipFile.Entries.Where(w =>
            w.FullName.TrimStart('/').StartsWith(ExcelFileNames.WorksheetBase, StringComparison.OrdinalIgnoreCase)
        ).ToList();

        foreach (var sheet in sheets)
        {
            // XRowInfos musy be cleared for every sheet or it'll cause duplicates
            _xRowInfos.Clear();
            _xMergeCellInfos.Clear();
            _newXMergeCellInfos.Clear();
            _calcChainCellRefs.Clear();

            var entry = archive.ZipFile.CreateEntry(sheet.FullName);

            var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

            var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableZipStream = zipStream.ConfigureAwait(false);

            await GenerateSheetByUpdateModeAsync(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true, cancellationToken).ConfigureAwait(false);
        }
    }
}
