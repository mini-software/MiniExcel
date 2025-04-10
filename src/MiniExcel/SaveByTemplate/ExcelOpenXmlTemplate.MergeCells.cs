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
        public void MergeSameCells(string path)
        {
            using (var stream = FileHelper.OpenSharedRead(path))
                MergeSameCellsImpl(stream);
        }

        public void MergeSameCells(byte[] fileInBytes)
        {
            using (Stream stream = new MemoryStream(fileInBytes))
                MergeSameCellsImpl(stream);
        }

        private void MergeSameCellsImpl(Stream stream)
        {
            stream.CopyTo(_outputFileStream);

            var reader = new ExcelOpenXmlSheetReader(_outputFileStream, null);
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
                    GenerateSheetXmlImplByUpdateMode(sheet, zipStream, sheetStream, new Dictionary<string, object>(), sharedStrings, mergeCells: true);
                    //doc.Save(zipStream); //don't do it beacause: https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png
                }
            }

            archive.zipFile.Dispose();
        }

        public Task MergeSameCellsAsync(string path, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => MergeSameCells(path), cancellationToken);
        }

        public Task MergeSameCellsAsync(byte[] fileInBytes, CancellationToken cancellationToken = default(CancellationToken))
        {
            return Task.Run(() => MergeSameCells(fileInBytes), cancellationToken);
        }
    }
}