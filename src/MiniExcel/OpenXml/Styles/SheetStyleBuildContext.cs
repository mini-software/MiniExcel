using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles
{
    internal class SheetStyleBuildContext : IDisposable
    {
        private static readonly string _emptyStylesXml = ExcelOpenXmlUtils.MinifyXml
        ($@"
            <?xml version=""1.0"" encoding=""utf-8""?>
            <x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">                
            </x:styleSheet>"
        );

        private readonly Dictionary<string, ZipPackageInfo> _zipDictionary;
        private readonly MiniExcelZipArchive _archive;
        private readonly Encoding _encoding;
        private readonly ICollection<ExcelColumnAttribute> _columns;

        private StringReader emptyStylesXmlStringReader;
        private ZipArchiveEntry oldStyleXmlZipEntry;
        private ZipArchiveEntry newStyleXmlZipEntry;
        private Stream oldXmlReaderStream;
        private Stream newXmlWriterStream;
        private bool initialized = false;
        private bool finalized = false;
        private bool disposed = false;

        public SheetStyleBuildContext(Dictionary<string, ZipPackageInfo> zipDictionary, MiniExcelZipArchive archive, Encoding encoding, ICollection<ExcelColumnAttribute> columns)
        {
            _zipDictionary = zipDictionary;
            _archive = archive;
            _encoding = encoding;
            _columns = columns;
        }

        public XmlReader OldXmlReader { get; private set; }

        public XmlWriter NewXmlWriter { get; private set; }

        public SheetStyleElementInfos OldElementInfos { get; private set; }

        public SheetStyleElementInfos GenerateElementInfos { get; private set; }

        public IEnumerable<ExcelColumnAttribute> ColumnsToApply { get; private set; }

        public int CustomFormatCount { get; private set; }

        public void Initialize(SheetStyleElementInfos generateElementInfos)
        {
            if (initialized)
            {
                throw new InvalidOperationException("The context has been initialized.");
            }

            oldStyleXmlZipEntry = _archive.Mode == ZipArchiveMode.Update ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) : null;
            if (oldStyleXmlZipEntry != null)
            {
                using (var oldStyleXmlStream = oldStyleXmlZipEntry.Open())
                {
                    OldElementInfos = ReadSheetStyleElementInfos(XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings() { IgnoreWhitespace = true }));
                }

                oldXmlReaderStream = oldStyleXmlZipEntry.Open();
                OldXmlReader = XmlReader.Create(oldXmlReaderStream, new XmlReaderSettings() { IgnoreWhitespace = true });

                newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
            }
            else
            {
                OldElementInfos = new SheetStyleElementInfos();

                emptyStylesXmlStringReader = new StringReader(_emptyStylesXml);
                OldXmlReader = XmlReader.Create(emptyStylesXmlStringReader, new XmlReaderSettings() { IgnoreWhitespace = true });

                newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
            }

            newXmlWriterStream = newStyleXmlZipEntry.Open();
            NewXmlWriter = XmlWriter.Create(newXmlWriterStream, new XmlWriterSettings() { Indent = true, Encoding = _encoding });

            GenerateElementInfos = generateElementInfos;
            ColumnsToApply = SheetStyleBuilderHelper.GenerateStyleIds(OldElementInfos.CellXfCount + generateElementInfos.CellXfCount, _columns).ToArray();//这里暂时加ToArray，避免多次计算，如果有性能问题再考虑优化
            CustomFormatCount = ColumnsToApply.Count();

            initialized = true;
        }

        public async Task InitializeAsync(SheetStyleElementInfos generateElementInfos)
        {
            if (initialized)
            {
                throw new InvalidOperationException("The context has been initialized.");
            }

            oldStyleXmlZipEntry = _archive.Mode == ZipArchiveMode.Update ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) : null;
            if (oldStyleXmlZipEntry != null)
            {
                using (var oldStyleXmlStream = oldStyleXmlZipEntry.Open())
                {
                    OldElementInfos = await ReadSheetStyleElementInfosAsync(XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings() { IgnoreWhitespace = true, Async = true }));
                }
                oldXmlReaderStream = oldStyleXmlZipEntry.Open();
                OldXmlReader = XmlReader.Create(oldXmlReaderStream, new XmlReaderSettings() { IgnoreWhitespace = true, Async = true });

                newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
            }
            else
            {
                OldElementInfos = new SheetStyleElementInfos();
                emptyStylesXmlStringReader = new StringReader(_emptyStylesXml);
                OldXmlReader = XmlReader.Create(emptyStylesXmlStringReader, new XmlReaderSettings() { IgnoreWhitespace = true, Async = true });

                newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
            }

            newXmlWriterStream = newStyleXmlZipEntry.Open();
            NewXmlWriter = XmlWriter.Create(newXmlWriterStream, new XmlWriterSettings() { Indent = true, Encoding = _encoding, Async = true });

            GenerateElementInfos = generateElementInfos;
            ColumnsToApply = SheetStyleBuilderHelper.GenerateStyleIds(OldElementInfos.CellXfCount + generateElementInfos.CellXfCount, _columns);
            CustomFormatCount = ColumnsToApply.Count();

            initialized = true;
        }

        public void FinalizeAndUpdateZipDictionary()
        {
            if (!initialized)
            {
                throw new InvalidOperationException("The context has not been initialized.");
            }
            if (disposed)
            {
                throw new ObjectDisposedException(nameof(SheetStyleBuildContext));
            }
            if (finalized)
            {
                throw new InvalidOperationException("The context has been finalized.");
            }
            try
            {
                OldXmlReader.Dispose();
                OldXmlReader = null;
                oldXmlReaderStream?.Dispose();
                oldXmlReaderStream = null;
                emptyStylesXmlStringReader?.Dispose();
                emptyStylesXmlStringReader = null;

                NewXmlWriter.Flush();
                NewXmlWriter.Close();
                NewXmlWriter.Dispose();
                NewXmlWriter = null;
                newXmlWriterStream.Dispose();
                newXmlWriterStream = null;

                if (oldStyleXmlZipEntry == null)
                {
                    _zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(newStyleXmlZipEntry, ExcelContentTypes.Styles));
                }
                else
                {
                    oldStyleXmlZipEntry?.Delete();
                    oldStyleXmlZipEntry = null;
                    var finalStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
                    using (var tempStream = newStyleXmlZipEntry.Open())
                    using (var newStream = finalStyleXmlZipEntry.Open())
                    {
                        tempStream.CopyTo(newStream);
                    }
                    _zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
                    newStyleXmlZipEntry.Delete();
                    newStyleXmlZipEntry = null;
                }

                finalized = true;
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to finalize and replace styles.", ex);
            }
        }

        public async Task FinalizeAndUpdateZipDictionaryAsync()
        {
            if (!initialized)
            {
                throw new InvalidOperationException("The context has not been initialized.");
            }
            if (disposed)
            {
                throw new ObjectDisposedException(nameof(SheetStyleBuildContext));
            }
            if (finalized)
            {
                throw new InvalidOperationException("The context has been finalized.");
            }
            try
            {
                OldXmlReader.Dispose();
                OldXmlReader = null;
                oldXmlReaderStream?.Dispose();
                oldXmlReaderStream = null;
                emptyStylesXmlStringReader?.Dispose();
                emptyStylesXmlStringReader = null;

                await NewXmlWriter.FlushAsync();
                NewXmlWriter.Close();
                NewXmlWriter.Dispose();
                NewXmlWriter = null;
                newXmlWriterStream.Dispose();
                newXmlWriterStream = null;

                if (oldStyleXmlZipEntry == null)
                {
                    _zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(newStyleXmlZipEntry, ExcelContentTypes.Styles));
                }
                else
                {
                    oldStyleXmlZipEntry?.Delete();
                    oldStyleXmlZipEntry = null;
                    var finalStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
                    using (var tempStream = newStyleXmlZipEntry.Open())
                    using (var newStream = finalStyleXmlZipEntry.Open())
                    {
                        await tempStream.CopyToAsync(newStream);
                    }
                    _zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
                    newStyleXmlZipEntry.Delete();
                    newStyleXmlZipEntry = null;
                }

                finalized = true;
            }
            catch (Exception ex)
            {
                throw new Exception("Failed to finalize and replace styles.", ex);
            }
        }

        private static SheetStyleElementInfos ReadSheetStyleElementInfos(XmlReader reader)
        {
            var elementInfos = new SheetStyleElementInfos();
            while (reader.Read())
            {
                SetElementInfos(reader, elementInfos);
            }
            return elementInfos;
        }

        private static async Task<SheetStyleElementInfos> ReadSheetStyleElementInfosAsync(XmlReader reader)
        {
            var elementInfos = new SheetStyleElementInfos();
            while (await reader.ReadAsync())
            {
                SetElementInfos(reader, elementInfos);
            }
            return elementInfos;
        }

        private static void SetElementInfos(XmlReader reader, SheetStyleElementInfos elementInfos)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.LocalName)
                {
                    case "numFmts":
                        elementInfos.ExistsNumFmts = true;
                        elementInfos.NumFmtCount = GetCount();
                        break;
                    case "fonts":
                        elementInfos.ExistsFonts = true;
                        elementInfos.FontCount = GetCount();
                        break;
                    case "fills":
                        elementInfos.ExistsFills = true;
                        elementInfos.FillCount = GetCount();
                        break;
                    case "borders":
                        elementInfos.ExistsBorders = true;
                        elementInfos.BorderCount = GetCount();
                        break;
                    case "cellStyleXfs":
                        elementInfos.ExistsCellStyleXfs = true;
                        elementInfos.CellStyleXfCount = GetCount();
                        break;
                    case "cellXfs":
                        elementInfos.ExistsCellXfs = true;
                        elementInfos.CellXfCount = GetCount();
                        break;
                }

                int GetCount()
                {
                    string count = reader.GetAttribute("count");
                    if (!string.IsNullOrEmpty(count) && int.TryParse(count, out int countValue))
                    {
                        return countValue;
                    }
                    return 0;
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    OldXmlReader?.Dispose();
                    oldXmlReaderStream?.Dispose();
                    emptyStylesXmlStringReader?.Dispose();

                    NewXmlWriter?.Dispose();
                    newXmlWriterStream?.Dispose();
                }

                disposed = true;
            }
        }
    }
}
