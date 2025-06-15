using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml.SaveByTemplate
{
    internal partial class ExcelOpenXmlTemplate
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task MergeSameCellsAsync(string path, CancellationToken ct = default)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                await MergeSameCellsImplAsync(stream, ct).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        public async Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken ct = default)
        {
            using (Stream stream = new MemoryStream(fileInBytes))
                await MergeSameCellsImplAsync(stream, ct).ConfigureAwait(false);
        }

        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        private async Task MergeSameCellsImplAsync(Stream stream, CancellationToken ct = default)
        {
            await stream.CopyToAsync(_outputFileStream
#if NETCOREAPP2_1_OR_GREATER
                                    , ct
#endif
                ).ConfigureAwait(false);

            var reader = await ExcelOpenXmlSheetReader.CreateAsync (_outputFileStream, null, ct: ct).ConfigureAwait(false);
            var archive = new ExcelOpenXmlZip(_outputFileStream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);

            //read sharedString
            var sharedStrings = reader._sharedStrings;

            //read all xlsx sheets
            var sheets = archive.zipFile.Entries.Where(w =>
                w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
            ).ToList();

            foreach (var sheet in sheets)
            {
                _xRowInfos = new List<XRowInfo>(); //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
                _xMergeCellInfos = new Dictionary<string, XMergeCell>();
                _newXMergeCellInfos = new List<XMergeCell>();

                var sheetStream = sheet.Open();
                var fullName = sheet.FullName;

                var entry = archive.zipFile.CreateEntry(fullName);
                using (var zipStream = entry.Open())
                {
                    await GenerateSheetXmlImplByUpdateModeAsync(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true, ct).ConfigureAwait(false);
                    //doc.Save(zipStream); //don't do it beacause: https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png
                }
            }

            archive.zipFile.Dispose();
        }
    }
}