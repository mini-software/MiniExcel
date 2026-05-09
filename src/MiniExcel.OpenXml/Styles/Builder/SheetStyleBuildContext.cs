using MiniExcelLib.OpenXml.Constants;

namespace MiniExcelLib.OpenXml.Styles.Builder;

internal sealed partial class SheetStyleBuildContext(Dictionary<string, ZipPackageInfo> zipDictionary, ZipArchive archive, Encoding encoding) : IDisposable, IAsyncDisposable
{
    private const string EmptyStylesXml = 
        """
        <?xml version="1.0" encoding="utf-8"?>
        <x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """;

    private readonly Dictionary<string, ZipPackageInfo> _zipDictionary = zipDictionary;
    private readonly ZipArchive _archive = archive;
    private readonly Encoding _encoding = encoding;

    private StringReader? _emptyStylesXmlStringReader;
    private ZipArchiveEntry? _oldStyleXmlZipEntry;
    private ZipArchiveEntry? _newStyleXmlZipEntry;
    private Stream? _oldXmlReaderStream;
    private Stream? _newXmlWriterStream;

    private bool _initialized;
    private bool _finalized;
    private bool _disposed;

    internal readonly SheetStyleFormatsCache SheetStyleFormatsCache = new();

    public XmlReader? OldXmlReader { get; private set; }
    public XmlWriter? NewXmlWriter { get; private set; }
    public SheetStyleElementInfos OldElementInfos { get; private set; } = null!;
    public SheetStyleElementInfos GeneratedElementInfos { get; private set; } = null!;
    public int CustomFormatCount => SheetStyleFormatsCache.FormatMappingsCount;

    [CreateSyncVersion]
    public async Task CreateAsync(SheetStyleElementInfos generatedElementInfos, CancellationToken cancellationToken = default)
    {
        const bool isAsync =
#if SYNC_ONLY
            false;
#else
            true;
#endif

        SheetStyleElementInfos infos;
        var styleEntry = _archive.Mode == ZipArchiveMode.Update
            ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles)
            : null;

        if (styleEntry is not null)
        {
#if NET8_0_OR_GREATER
            var oldStyleXmlStream = await styleEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableStream = oldStyleXmlStream.ConfigureAwait(false);
#else
            using var oldStyleXmlStream = styleEntry.Open();
#endif
            using var reader = XmlReader.Create(oldStyleXmlStream, XmlReaderHelper.GetXmlReaderSettings(isAsync));
            infos = await ReadSheetStyleElementInfosAsync(reader, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            infos = new SheetStyleElementInfos();
        }

        SheetStyleFormatsCache.SetCurrentIndex(infos.CellXfCount + generatedElementInfos.CellXfCount);
    }

    [CreateSyncVersion]
    public async Task InitializeAsync(SheetStyleElementInfos generatedElementInfos, CancellationToken cancellationToken = default)
    {
        if (_initialized)
            throw new InvalidOperationException("The context has already been initialized.");

        const bool isAsync =
#if SYNC_ONLY
            false;
#else
            true;
#endif

        GeneratedElementInfos = generatedElementInfos;

        _oldStyleXmlZipEntry = _archive.Mode == ZipArchiveMode.Update
            ? _archive.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.Styles)
            : null;

        var xmlReaderSettings = XmlReaderHelper.GetXmlReaderSettings(isAsync);
        if (_oldStyleXmlZipEntry is not null)
        {
#if NET8_0_OR_GREATER
            var oldStyleXmlStream = await _oldStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using (_ = oldStyleXmlStream.ConfigureAwait(false))
#else
            using (var oldStyleXmlStream = _oldStyleXmlZipEntry.Open())
#endif
            {
                using var reader = XmlReader.Create(oldStyleXmlStream, xmlReaderSettings);
                OldElementInfos = await ReadSheetStyleElementInfosAsync(reader, cancellationToken).ConfigureAwait(false);
            }

#if NET8_0_OR_GREATER
            _oldXmlReaderStream = await _oldStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
            _oldXmlReaderStream = _oldStyleXmlZipEntry.Open();
#endif
            OldXmlReader = XmlReader.Create(_oldXmlReaderStream, xmlReaderSettings);
            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles + ".temp", CompressionLevel.Fastest);
        }
        else
        {
            OldElementInfos = new SheetStyleElementInfos();
            _emptyStylesXmlStringReader = new StringReader(EmptyStylesXml);
            OldXmlReader = XmlReader.Create(_emptyStylesXmlStringReader, xmlReaderSettings);

            _newStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
        }

#if NET8_0_OR_GREATER
        _newXmlWriterStream = await _newStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#else
        _newXmlWriterStream = _newStyleXmlZipEntry.Open();
#endif
        NewXmlWriter = XmlWriter.Create(_newXmlWriterStream, new XmlWriterSettings { Indent = true, Encoding = _encoding, Async = isAsync });

        _initialized = true;
    }
    
    public void UpdateFormatIds(ICollection<MiniExcelColumnMapping> mappings)
    {
        SheetStyleFormatsCache.AddMappings(mappings);
    }
    
    [CreateSyncVersion]
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
            OldXmlReader?.Dispose();
            OldXmlReader = null;
#if NET8_0_OR_GREATER
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

            await NewXmlWriter!.FlushAsync().ConfigureAwait(false);
#if NET8_0_OR_GREATER
            await NewXmlWriter.DisposeAsync().ConfigureAwait(false);
#else
            NewXmlWriter.Dispose();
#endif

            NewXmlWriter = null;

#if NET8_0_OR_GREATER
            await _newXmlWriterStream!.DisposeAsync().ConfigureAwait(false);
#else
            _newXmlWriterStream?.Dispose();
#endif
            _newXmlWriterStream = null;

            if (_oldStyleXmlZipEntry is null)
            {
                _zipDictionary.Add(ExcelFileNames.Styles, new ZipPackageInfo(_newStyleXmlZipEntry!, ExcelContentTypes.Styles));
            }
            else
            {
                _oldStyleXmlZipEntry?.Delete();
                _oldStyleXmlZipEntry = null;
                var finalStyleXmlZipEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);

#if NET8_0_OR_GREATER
                var tempStream = await _newStyleXmlZipEntry!.OpenAsync(cancellationToken).ConfigureAwait(false);
                var newStream = await finalStyleXmlZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                await using (_ = tempStream.ConfigureAwait(false))
                await using (_= newStream.ConfigureAwait(false))
#else
                using (var tempStream = _newStyleXmlZipEntry!.Open())
                using (var newStream = finalStyleXmlZipEntry.Open())
#endif
                {
                    await tempStream.CopyToAsync(newStream, 4096, cancellationToken).ConfigureAwait(false);
                }

                _zipDictionary[ExcelFileNames.Styles] = new ZipPackageInfo(finalStyleXmlZipEntry, ExcelContentTypes.Styles);
                _newStyleXmlZipEntry?.Delete();
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
        return;

        int GetCount()
        {
            var count = reader.GetAttribute("count");
            return int.TryParse(count, out var countValue) ? countValue : 0;
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        OldXmlReader?.Dispose();
        _oldXmlReaderStream?.Dispose();
        _emptyStylesXmlStringReader?.Dispose();

        NewXmlWriter?.Dispose();
        _newXmlWriterStream?.Dispose();

        _disposed = true;
    }

    public async ValueTask DisposeAsync()
    {
        if (_disposed)
            return;

        await CastAndDispose(_emptyStylesXmlStringReader).ConfigureAwait(false);
        await CastAndDispose(_oldXmlReaderStream).ConfigureAwait(false);
        await CastAndDispose(_newXmlWriterStream).ConfigureAwait(false);
        await CastAndDispose(OldXmlReader).ConfigureAwait(false);
        await CastAndDispose(NewXmlWriter).ConfigureAwait(false);

        _disposed = true;
        return;
    
        static async ValueTask CastAndDispose(IDisposable? resource)
        {
            if (resource is IAsyncDisposable asyncDisposable)
                await asyncDisposable.DisposeAsync().ConfigureAwait(false);
            else
                resource?.Dispose();
        }
    }
}
