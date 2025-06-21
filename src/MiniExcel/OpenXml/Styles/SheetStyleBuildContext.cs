using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles;

internal class SheetStyleBuildContext : IDisposable
{
    private static readonly string _emptyStylesXml = ExcelOpenXmlUtils.MinifyXml
    (@"
            <?xml version=""1.0"" encoding=""utf-8""?>
            <x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">                
            </x:styleSheet>"
    );

    private readonly Dictionary<string, ZipPackageInfo> _zipDictionary;
    private readonly MiniExcelZipArchive _archive;
    private readonly Encoding _encoding;
    private readonly ICollection<ExcelColumnAttribute> _columns;

    private StringReader _emptyStylesXmlStringReader;
    private ZipArchiveEntry? _oldStyleXmlZipEntry;
    private ZipArchiveEntry? _newStyleXmlZipEntry;
    private Stream _oldXmlReaderStream;
    private Stream _newXmlWriterStream;

    private bool _initialized;
    private bool _finalized;
    private bool _disposed;

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
        if (_initialized)
            throw new InvalidOperationException("The context has been initialized.");

        _oldStyleXmlZipEntry = _archive.Mode == ZipArchiveMode.Update ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) : null;
        if (_oldStyleXmlZipEntry is not null)
        {
            using (var oldStyleXmlStream = _oldStyleXmlZipEntry.Open())
            {
                using XmlReader reader = XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true });
                OldElementInfos = ReadSheetStyleElementInfos(reader);
            }

            _oldXmlReaderStream = _oldStyleXmlZipEntry.Open();
            OldXmlReader = XmlReader.Create(_oldXmlReaderStream, new XmlReaderSettings { IgnoreWhitespace = true });

            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
        }
        else
        {
            OldElementInfos = new SheetStyleElementInfos();

            _emptyStylesXmlStringReader = new StringReader(_emptyStylesXml);
            OldXmlReader = XmlReader.Create(_emptyStylesXmlStringReader, new XmlReaderSettings { IgnoreWhitespace = true });

            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
        }

        _newXmlWriterStream = _newStyleXmlZipEntry.Open();
        NewXmlWriter = XmlWriter.Create(_newXmlWriterStream, new XmlWriterSettings { Indent = true, Encoding = _encoding });

        GenerateElementInfos = generateElementInfos;
        ColumnsToApply = SheetStyleBuilderHelper.GenerateStyleIds(OldElementInfos.CellXfCount + generateElementInfos.CellXfCount, _columns).ToArray();//这里暂时加ToArray，避免多次计算，如果有性能问题再考虑优化
        CustomFormatCount = ColumnsToApply.Count();

        _initialized = true;
    }

    public async Task InitializeAsync(SheetStyleElementInfos generateElementInfos, CancellationToken cancellationToken = default)
    {
        if (_initialized)
            throw new InvalidOperationException("The context has already been initialized.");

        cancellationToken.ThrowIfCancellationRequested();

        _oldStyleXmlZipEntry = _archive.Mode == ZipArchiveMode.Update ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) : null;
        if (_oldStyleXmlZipEntry is not null)
        {
            using (var oldStyleXmlStream = _oldStyleXmlZipEntry.Open())
            {
                using var reader = XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });
                OldElementInfos = await ReadSheetStyleElementInfosAsync(reader, cancellationToken).ConfigureAwait(false);
            }
            _oldXmlReaderStream = _oldStyleXmlZipEntry.Open();
            OldXmlReader = XmlReader.Create(_oldXmlReaderStream, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });

            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
        }
        else
        {
            OldElementInfos = new SheetStyleElementInfos();
            _emptyStylesXmlStringReader = new StringReader(_emptyStylesXml);
            OldXmlReader = XmlReader.Create(_emptyStylesXmlStringReader, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });

            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
        }

        _newXmlWriterStream = _newStyleXmlZipEntry.Open();
        NewXmlWriter = XmlWriter.Create(_newXmlWriterStream, new XmlWriterSettings { Indent = true, Encoding = _encoding, Async = true });

        GenerateElementInfos = generateElementInfos;
        ColumnsToApply = SheetStyleBuilderHelper.GenerateStyleIds(OldElementInfos.CellXfCount + generateElementInfos.CellXfCount, _columns).ToArray();//ToArray to avoid multiple calculations, if there is a performance problem then consider optimizing the
        CustomFormatCount = ColumnsToApply.Count();

        _initialized = true;
    }

    public void FinalizeAndUpdateZipDictionary()
    {
        if (!_initialized)
            throw new InvalidOperationException("The context has not been initialized.");
        if (_disposed)
            throw new ObjectDisposedException(nameof(SheetStyleBuildContext));
        if (_finalized)
            throw new InvalidOperationException("The context has been finalized.");

        try
        {
            OldXmlReader.Dispose();
            OldXmlReader = null;
            _oldXmlReaderStream?.Dispose();
            _oldXmlReaderStream = null;

            _emptyStylesXmlStringReader?.Dispose();
            _emptyStylesXmlStringReader = null;

            NewXmlWriter.Flush();
            NewXmlWriter.Close();
            NewXmlWriter.Dispose();
            NewXmlWriter = null;

            _newXmlWriterStream.Dispose();
            _newXmlWriterStream = null;

            if (_oldStyleXmlZipEntry is null)
            {
                _zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(_newStyleXmlZipEntry, ExcelContentTypes.Styles));
            }
            else
            {
                _oldStyleXmlZipEntry?.Delete();
                _oldStyleXmlZipEntry = null;
                var finalStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);

                using (var tempStream = _newStyleXmlZipEntry.Open())
                using (var newStream = finalStyleXmlZipEntry.Open())
                {
                    tempStream.CopyTo(newStream);
                }

                _zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
                _newStyleXmlZipEntry.Delete();
                _newStyleXmlZipEntry = null;
            }

            _finalized = true;
        }
        catch (Exception ex)
        {
            throw new Exception("Failed to finalize and replace styles.", ex);
        }
    }

    public async Task FinalizeAndUpdateZipDictionaryAsync(CancellationToken cancellationToken = default)
    {
        if (!_initialized)
            throw new InvalidOperationException("The context has not been initialized.");
        if (_disposed)
            throw new ObjectDisposedException(nameof(SheetStyleBuildContext));
        if (_finalized)
            throw new InvalidOperationException("The context has been finalized.");

        try
        {
            cancellationToken.ThrowIfCancellationRequested();

            OldXmlReader.Dispose();
            OldXmlReader = null;
#if NET5_0_OR_GREATER
            if (_oldXmlReaderStream is not null)
            {
                await _oldXmlReaderStream.DisposeAsync().ConfigureAwait(false);
            }
#else
                _oldXmlReaderStream?.Dispose();
#endif
            _oldXmlReaderStream = null;

            _emptyStylesXmlStringReader?.Dispose();
            _emptyStylesXmlStringReader = null;

            await NewXmlWriter.FlushAsync().ConfigureAwait(false);
            NewXmlWriter.Close();
            //NewXmlWriter.Dispose();

#if NET5_0_OR_GREATER
            await NewXmlWriter.DisposeAsync().ConfigureAwait(false);
#else
                NewXmlWriter.Dispose();
#endif

            NewXmlWriter = null;

#if NET5_0_OR_GREATER
            await _newXmlWriterStream.DisposeAsync().ConfigureAwait(false);
#else
                _newXmlWriterStream.Dispose();
#endif
            _newXmlWriterStream = null;

            if (_oldStyleXmlZipEntry is null)
            {
                _zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(_newStyleXmlZipEntry, ExcelContentTypes.Styles));
            }
            else
            {
                _oldStyleXmlZipEntry?.Delete();
                _oldStyleXmlZipEntry = null;
                var finalStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);

                using (var tempStream = _newStyleXmlZipEntry.Open())
                using (var newStream = finalStyleXmlZipEntry.Open())
                {
                    await tempStream.CopyToAsync(newStream, 4096, cancellationToken).ConfigureAwait(false);
                }

                _zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
                _newStyleXmlZipEntry.Delete();
                _newStyleXmlZipEntry = null;
            }

            _finalized = true;
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

    private static async Task<SheetStyleElementInfos> ReadSheetStyleElementInfosAsync(XmlReader reader, CancellationToken cancellationToken = default)
    {
        var elementInfos = new SheetStyleElementInfos();
        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            SetElementInfos(reader, elementInfos);
        }
        return elementInfos;
    }

    private static void SetElementInfos(XmlReader reader, SheetStyleElementInfos elementInfos)
    {
        if (reader.NodeType != XmlNodeType.Element)
            return;

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
            var count = reader.GetAttribute("count") ?? string.Empty;
            return int.TryParse(count, out var countValue) ? countValue : 0;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            OldXmlReader?.Dispose();
            _oldXmlReaderStream?.Dispose();
            _emptyStylesXmlStringReader?.Dispose();

            NewXmlWriter?.Dispose();
            _newXmlWriterStream?.Dispose();
        }

        _disposed = true;
    }
}