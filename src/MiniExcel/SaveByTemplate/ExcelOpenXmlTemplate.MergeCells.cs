using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MiniExcelLibs.OpenXml;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;

namespace MiniExcelLibs.SaveByTemplate;

internal partial class ExcelOpenXmlTemplate
{
    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default)
    {
        using var stream = FileHelper.OpenSharedRead(path);
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public async Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default)
    {
        using Stream stream = new MemoryStream(fileInBytes);
        await MergeSameCellsImplAsync(stream, cancellationToken).ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    private async Task MergeSameCellsImplAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        await stream.CopyToAsync(_outputFileStream
#if NETCOREAPP2_1_OR_GREATER
            , cancellationToken
#endif
        ).ConfigureAwait(false);

        using var reader = await ExcelOpenXmlSheetReader.CreateAsync(_outputFileStream, null, cancellationToken: cancellationToken).ConfigureAwait(false);
        using var archive = new ExcelOpenXmlZip(_outputFileStream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);

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

            var sheetStream = sheet.Open();
            var fullName = sheet.FullName;

            var entry = archive.ZipFile.CreateEntry(fullName);
            using var zipStream = entry.Open();
            await GenerateSheetXmlImplByUpdateModeAsync(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true, cancellationToken).ConfigureAwait(false);
            //doc.Save(zipStream); //don't do it beacause: https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png
        }

        archive.ZipFile.Dispose();
    }
}