using MiniExcelLibs.Attributes;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Zip;
using System.IO.Compression;
using System.Text;
using System.Xml;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.OpenXml.Styles;

internal sealed class SheetStyleBuildContext(Dictionary<string, ZipPackageInfo> zipDictionary, MiniExcelZipArchive archive, Encoding encoding) : IDisposable
{
    private const string EmptyStylesXml = 
        """
        <?xml version="1.0" encoding="utf-8"?>
        <x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />                
        """;

    private StringReader _emptyStylesXmlStringReader;
    private ZipArchiveEntry _oldStyleXmlZipEntry;
    private ZipArchiveEntry _newStyleXmlZipEntry;
    private Stream _oldXmlReaderStream;
    private Stream _newXmlWriterStream;
        
    private bool _initialized;
    private bool _finalized;
    private bool _disposed;
    
    internal readonly SheetStyleFormatsCache SheetStyleFormatsCache = new();

    public XmlReader OldXmlReader { get; private set; }
    public XmlWriter NewXmlWriter { get; private set; }
    public SheetStyleElementInfos OldElementInfos { get; private set; }
    public SheetStyleElementInfos GenerateElementInfos { get; private set; }
    public int CustomFormatCount => SheetStyleFormatsCache.FormatMappingsCount;

    public void Create(SheetStyleElementInfos generatedElementInfos)
    {
        SheetStyleElementInfos infos;
        var styleEntry = archive.Mode == ZipArchiveMode.Update
            ? archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles)
            : null;

        if (styleEntry is not null)
        {
            using var oldStyleXmlStream = styleEntry.Open();
            using var reader = XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true });
            infos = ReadSheetStyleElementInfos(reader);
        }
        else
        {
            infos = new SheetStyleElementInfos();
        }

        SheetStyleFormatsCache.SetCurrentIndex(infos.CellXfCount + generatedElementInfos.CellXfCount);
    }

    public async Task CreateAsync(SheetStyleElementInfos generatedElementInfos, CancellationToken cancellationToken = default)
    {
        SheetStyleElementInfos infos;
        var styleEntry = archive.Mode == ZipArchiveMode.Update
            ? archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles)
            : null;

        if (styleEntry is not null)
        {
#if NET10_0_OR_GREATER
            var oldStyleXmlStream = await styleEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableStream = oldStyleXmlStream.ConfigureAwait(false);
#else
            using var oldStyleXmlStream = styleEntry.Open();
#endif
            using var reader = XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });
            infos = await ReadSheetStyleElementInfosAsync(reader, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            infos = new SheetStyleElementInfos();
        }

        SheetStyleFormatsCache.SetCurrentIndex(infos.CellXfCount + generatedElementInfos.CellXfCount);
    }
    
    public void Initialize(SheetStyleElementInfos generateElementInfos)
    {
        if (_initialized)
            throw new InvalidOperationException("The context has been initialized.");

        GenerateElementInfos =  generateElementInfos; 
        
        _oldStyleXmlZipEntry = archive.Mode == ZipArchiveMode.Update 
            ? archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) 
            : null;

        if (_oldStyleXmlZipEntry != null)
        {
            using (var oldStyleXmlStream = _oldStyleXmlZipEntry.Open())
            {
                OldElementInfos = ReadSheetStyleElementInfos(XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true }));
            }

            _oldXmlReaderStream = _oldStyleXmlZipEntry.Open();
            OldXmlReader = XmlReader.Create(_oldXmlReaderStream, new XmlReaderSettings { IgnoreWhitespace = true });

            _newStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
        }
        else
        {
            OldElementInfos = new SheetStyleElementInfos();

            _emptyStylesXmlStringReader = new StringReader(EmptyStylesXml);
            OldXmlReader = XmlReader.Create(_emptyStylesXmlStringReader, new XmlReaderSettings { IgnoreWhitespace = true });

            _newStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
        }

        _newXmlWriterStream = _newStyleXmlZipEntry.Open();
        NewXmlWriter = XmlWriter.Create(_newXmlWriterStream, new XmlWriterSettings { Indent = true, Encoding = encoding });

        _initialized = true;
    }

    public async Task InitializeAsync(SheetStyleElementInfos generateElementInfos, CancellationToken cancellationToken = default)
    {
        if (_initialized)
            throw new InvalidOperationException("The context has already been initialized.");

        GenerateElementInfos = generateElementInfos;

        _oldStyleXmlZipEntry = archive.Mode == ZipArchiveMode.Update 
            ? archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles) 
            : null;

        if (_oldStyleXmlZipEntry != null)
        {
#if NET10_0_OR_GREATER
            var oldStyleXmlStream = await _oldStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using (_ = oldStyleXmlStream.ConfigureAwait(false))
#else
            using (var oldStyleXmlStream = _oldStyleXmlZipEntry.Open())
#endif
            {
                OldElementInfos = await ReadSheetStyleElementInfosAsync(XmlReader.Create(oldStyleXmlStream, new XmlReaderSettings { IgnoreWhitespace = true, Async = true }), cancellationToken);
            }

#if NET10_0_OR_GREATER
            _oldXmlReaderStream = await _oldStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
            _oldXmlReaderStream = _oldStyleXmlZipEntry.Open();
#endif
            OldXmlReader = XmlReader.Create(_oldXmlReaderStream, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });
            _newStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
        }
        else
        {
            OldElementInfos = new SheetStyleElementInfos();
            _emptyStylesXmlStringReader = new StringReader(EmptyStylesXml);
            OldXmlReader = XmlReader.Create(_emptyStylesXmlStringReader, new XmlReaderSettings { IgnoreWhitespace = true, Async = true });

            _newStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
        }

#if NET10_0_OR_GREATER
        _newXmlWriterStream = await _newStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        _newXmlWriterStream = _newStyleXmlZipEntry.Open();
#endif
        NewXmlWriter = XmlWriter.Create(_newXmlWriterStream, new XmlWriterSettings { Indent = true, Encoding = encoding, Async = true });

        _initialized = true;
    }
    
    public void UpdateFormatIds(ICollection<ExcelColumnInfo> mappings)
    {
        SheetStyleFormatsCache.AddMappings(mappings);
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

            if (_oldStyleXmlZipEntry == null)
            {
                zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(_newStyleXmlZipEntry, ExcelContentTypes.Styles));
            }
            else
            {
                _oldStyleXmlZipEntry?.Delete();
                _oldStyleXmlZipEntry = null;
                var finalStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
                    
                using (var tempStream = _newStyleXmlZipEntry.Open())
                using (var newStream = finalStyleXmlZipEntry.Open())
                {
                    tempStream.CopyTo(newStream);
                }
                    
                zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
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
            _oldXmlReaderStream?.Dispose();
            _oldXmlReaderStream = null;

            _emptyStylesXmlStringReader?.Dispose();
            _emptyStylesXmlStringReader = null;

            await NewXmlWriter.FlushAsync();
            NewXmlWriter.Close();
            NewXmlWriter.Dispose();
            NewXmlWriter = null;
                
            _newXmlWriterStream.Dispose();
            _newXmlWriterStream = null;

            if (_oldStyleXmlZipEntry == null)
            {
                zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(_newStyleXmlZipEntry, ExcelContentTypes.Styles));
            }
            else
            {
                _oldStyleXmlZipEntry?.Delete();
                _oldStyleXmlZipEntry = null;
                var finalStyleXmlZipEntry = archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
                    
                using (var tempStream = _newStyleXmlZipEntry.Open())
                using (var newStream = finalStyleXmlZipEntry.Open())
                {
                    await tempStream.CopyToAsync(newStream, 4096, cancellationToken);
                }
                    
                zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
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
        while (await reader.ReadAsync())
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
    }

    private void Dispose(bool disposing)
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