namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default)
    {
#if NETSTANDARD2_0
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
#else
        var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await using var disposableStream = stream.ConfigureAwait(false);
#endif
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default)
    {
#if NETSTANDARD2_0
        using var stream = new MemoryStream(fileInBytes);
#else
        var stream = new MemoryStream(fileInBytes);
        await using var disposableStream = stream.ConfigureAwait(false);
#endif
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task MergeSameCellsImplAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        await stream.CopyToAsync(_outputFileStream
#if NET8_0_OR_GREATER
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
            w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
            w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        foreach (var sheet in sheets)
        {
            // XRowInfos musy be cleared for every sheet or it'll cause duplicates
            _xRowInfos.Clear();
            _xMergeCellInfos.Clear();
            _newXMergeCellInfos.Clear();
            _calcChainCellRefs.Clear();

            var entry = archive.ZipFile.CreateEntry(sheet.FullName);

#if NET8_0_OR_GREATER
            var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

            var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableZipStream = zipStream.ConfigureAwait(false);
#else
            using var sheetStream = sheet.Open();
            using var zipStream = entry.Open();
#endif

            await GenerateSheetByUpdateModeAsync(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true, cancellationToken).ConfigureAwait(false);
        }
    }
}
