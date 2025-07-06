namespace MiniExcelLib.Core.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    public async Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default)
    {
        using Stream stream = new MemoryStream(fileInBytes);
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task MergeSameCellsImplAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        await stream.CopyToAsync(_outputFileStream
#if NETCOREAPP2_1_OR_GREATER
            , cancellationToken
#endif
        ).ConfigureAwait(false);

        using var reader = await OpenXmlReader.CreateAsync(_outputFileStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        using var archive = new OpenXmlZip(_outputFileStream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);

        //read sharedString
        var sharedStrings = reader.SharedStrings;

        //read all xlsx sheets
        var sheets = archive.ZipFile.Entries.Where(w =>
            w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
            w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
        ).ToList();

        foreach (var sheet in sheets)
        {
            _xRowInfos = []; //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
            _xMergeCellInfos = [];
            _newXMergeCellInfos = [];

#if NET10_0_OR_GREATER
            var sheetStream = await sheet.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
            var sheetStream = sheet.Open();
#endif
            var fullName = sheet.FullName;

            var entry = archive.ZipFile.CreateEntry(fullName);
#if NET10_0_OR_GREATER
            using var zipStream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
            using var zipStream = entry.Open();
#endif
            await GenerateSheetXmlImplByUpdateModeAsync(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true, cancellationToken).ConfigureAwait(false);
            //doc.Save(zipStream); //don't do it beacause: https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png
        }

#if NET10_0_OR_GREATER
        await archive.ZipFile.DisposeAsync().ConfigureAwait(false);
#else
        archive.ZipFile.Dispose();
#endif
    }
}