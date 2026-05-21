using System.ComponentModel;
using System.Xml.Linq;
using MiniExcelLib.Core;
using MiniExcelLib.OpenXml.Constants;
using MiniExcelLib.OpenXml.Styles.Builder;

namespace MiniExcelLib.OpenXml.Writer;

internal partial class OpenXmlWriter
{
    private static readonly string[] EntriesToIgnoreOnCopy = [
        ExcelFileNames.ContentTypes, 
        ExcelFileNames.Workbook, 
        ExcelFileNames.WorkbookRels, 
        ExcelFileNames.SharedStrings, 
        ExcelFileNames.Styles
    ];

    private readonly Stream? _oldStream;
    private readonly ZipArchive? _oldArchive;


    private OpenXmlWriter(Stream oldStream, Stream newStream, ZipArchive oldArchive, ZipArchive newArchive, object? value, string sheetName, OpenXmlConfiguration configuration, bool printHeader)
        : this(newStream, newArchive, value, sheetName, configuration, printHeader) 
    {
        _oldStream = oldStream;
        _oldArchive = oldArchive;
    }

    [CreateSyncVersion]
    internal static async ValueTask<OpenXmlWriter> CreateForCopyAsync(Stream inStream, Stream outStream, object? value, string sheetName, bool printHeader, IMiniExcelConfiguration? configuration, CancellationToken cancellationToken = default)
    {
        ThrowHelper.ThrowIfInvalidSheetName(sheetName);

        var conf = configuration as OpenXmlConfiguration ?? OpenXmlConfiguration.Default;
        if (conf is { EnableAutoWidth: true, FastMode: false })
            throw new InvalidOperationException("Auto width requires fast mode to be enabled");

        var oldArchive = await ZipArchive.CreateAsync(inStream, ZipArchiveMode.Read, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);
        var newArchive = await ZipArchive.CreateAsync(outStream, ZipArchiveMode.Create, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);
        return new OpenXmlWriter(inStream, outStream, oldArchive, newArchive, value, sheetName, conf, printHeader);
    }

    [CreateSyncVersion]
    public async Task<int> CopyAndInsertAsync(bool overwriteSheet = false, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
#if NET10_0_OR_GREATER
        await using var disposableOldArchive = _oldArchive!.ConfigureAwait(false);
        await using var disposableNewArchive = _archive.ConfigureAwait(false);
#else
        using var disposableOldArchive = _oldArchive;
        using var disposableNewArchive = _archive;
#endif
        using var reader = await OpenXmlReader.CreateAsync(_oldStream!, _configuration, cancellationToken: cancellationToken).ConfigureAwait(false);
        var rels = await reader.GetWorkbookRelsAsync(_oldArchive!.Entries, cancellationToken).ConfigureAwait(false) ?? [];

        _sheets.AddRange(rels
            .OrderBy(sheet => sheet.Id)
            .Select(sheet => new SheetDto
            {
                Name = sheet.Name,
                SheetIdx = (int)sheet.Id,
                State = sheet.State
            })
        );

        var existingSheetDto = _sheets.SingleOrDefault(s => s.Name == _sheetName);
        if (existingSheetDto is not null && !overwriteSheet)
            throw new InvalidOperationException($"Sheet \"{_sheetName}\" already exists");

        var sheetStylesBuilderUtils = await CopySheetStylesAndGetBuilderUtilsAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableSheetStylesBuilderUtils = sheetStylesBuilderUtils.ConfigureAwait(false);

        await _sheetStyleBuilderContext.DisposeAsync().ConfigureAwait(false);
        _sheetStyleBuilderContext = sheetStylesBuilderUtils.SheetStyleBuilderContext;

        var sharedStringsEntry = _oldArchive.GetEntry(ExcelFileNames.SharedStrings);
        if (sharedStringsEntry is not null)
        {
            foreach (var (key, value) in reader.SharedStrings)
            {
                _sharedStrings[value] = key;
            }
        }

        int rowsWritten;
        List<string> entriesToIgnoreOnCopy = [..EntriesToIgnoreOnCopy];

        if (existingSheetDto is null)
        {
            _currentSheetIndex = (int)rels.Max(m => m.Id) + 1;
            var newSheetInfoDto = GetSheetInfos(_sheetName).ToDto(_currentSheetIndex);
            _sheets.Add(newSheetInfoDto);
         
            rowsWritten = await CreateSheetXmlAsync(_value, newSheetInfoDto.Path, progress, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            _currentSheetIndex = existingSheetDto.SheetIdx;
            rowsWritten = await CreateSheetXmlAsync(_value, existingSheetDto.Path, progress, cancellationToken).ConfigureAwait(false);
            entriesToIgnoreOnCopy.AddRange([
                ExcelFileNames.Worksheet(_currentSheetIndex),
                ExcelFileNames.SheetRels(_currentSheetIndex),
                ExcelFileNames.Drawing(_currentSheetIndex),
                ExcelFileNames.DrawingRels(_currentSheetIndex)
            ]);
        }

        foreach (var entry in _oldArchive.Entries.ExceptBy(entriesToIgnoreOnCopy, e => e.FullName, StringComparer.InvariantCultureIgnoreCase))
        {
            await CopyEntryAsync(entry, cancellationToken).ConfigureAwait(false);
        }

        await sheetStylesBuilderUtils.SheetStyleBuilder.BuildAsync(cancellationToken).ConfigureAwait(false);
        if (sheetStylesBuilderUtils.Archive.GetEntry(ExcelFileNames.Styles) is { } tempStylesEntry)
        {
            var newStylesEntry = _archive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
            var newStylesEntryStream = await newStylesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            var tempStylesEntryStream = await tempStylesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);

#if NET8_0_OR_GREATER
            await using var disposableNewStylesEntryStream = newStylesEntryStream.ConfigureAwait(false);
            await using var disposableTempStylesEntryStream = tempStylesEntryStream.ConfigureAwait(false);
#else
            using var disposableNewStylesEntryStream = newStylesEntryStream;
            using var disposableTempStylesEntryStream = tempStylesEntryStream;
#endif

            await tempStylesEntryStream.CopyToAsync(newStylesEntryStream, 81920, cancellationToken).ConfigureAwait(false);
            await newStylesEntryStream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        
        await AddFilesToZipAsync(cancellationToken).ConfigureAwait(false);
        await GenerateSharedStringsAsync(cancellationToken).ConfigureAwait(false);
        await GenerateDrawingRelXmlAsync(_currentSheetIndex, cancellationToken).ConfigureAwait(false);
        await GenerateDrawingXmlAsync(_currentSheetIndex, cancellationToken).ConfigureAwait(false);
        
        await CreateZipEntryAsync(
            ExcelFileNames.SheetRels(_currentSheetIndex),
            null,
            ExcelXml.DefaultSheetRelXml.Replace("{{format}}", ExcelXml.DrawingRelationship(_currentSheetIndex)),
            cancellationToken).ConfigureAwait(false);

        var (workbookXml, workbookRelsXml, _) = GenerateWorkbookXmls();
        await CreateZipEntryAsync(
            ExcelFileNames.Workbook,
            ExcelContentTypes.Workbook,
            ExcelXml.DefaultWorkbookXml.Replace("{{sheets}}", workbookXml),
            cancellationToken).ConfigureAwait(false);

        await CreateZipEntryAsync(
            ExcelFileNames.WorkbookRels,
            null,
            ExcelXml.DefaultWorkbookXmlRels.Replace("{{sheets}}", workbookRelsXml),
            cancellationToken).ConfigureAwait(false);
        
        await CopyAndUpdateContentTypesAsync(cancellationToken).ConfigureAwait(false);

        return rowsWritten;
    }
    
    [CreateSyncVersion]
    private async Task<TempSheetStylesBuilderUtils> CopySheetStylesAndGetBuilderUtilsAsync(CancellationToken cancellationToken = default)
    {
        var backingStream = new MemoryStream();
        var tempArchive = await ZipArchive.CreateAsync(backingStream, ZipArchiveMode.Create, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);
#if NET10_0_OR_GREATER
        await using (_ = tempArchive.ConfigureAwait(false))
#else
        using (_ = tempArchive)
#endif
        {
            var tempStylesEntry = tempArchive.CreateEntry(ExcelFileNames.Styles, CompressionLevel.Fastest);
            if (_oldArchive?.GetEntry(ExcelFileNames.Styles) is { } stylesEntry)
            {
                var oldEntryStream = await stylesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
                var tempEntryStream = await tempStylesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);

                await oldEntryStream.CopyToAsync(tempEntryStream, 81920, cancellationToken).ConfigureAwait(false);
                await tempEntryStream.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
        }

        backingStream.Seek(0, SeekOrigin.Begin);
        var copiedArchive = await ZipArchive.CreateAsync(backingStream, ZipArchiveMode.Update, true, Utf8WithBom, cancellationToken).ConfigureAwait(false);

        SheetStyleBuilderContext? oldStylesContext = null;
        try
        {
            oldStylesContext = new SheetStyleBuilderContext(_zipContentsMap, copiedArchive, Utf8WithBom);
            SheetStyleBuilderBase builder = _configuration.TableStyles switch
            {
                TableStyles.None => new MinimalSheetStyleBuilder(oldStylesContext),
                TableStyles.Default => new DefaultSheetStyleBuilder(oldStylesContext, _configuration.StyleOptions),
                _ => throw new InvalidEnumArgumentException(nameof(_configuration.TableStyles), (int)_configuration.TableStyles, typeof(TableStyles))
            };

            var newInfos = builder.GetGeneratedElementInfos();
            await oldStylesContext.CreateAsync(newInfos, cancellationToken).ConfigureAwait(false);

            var copiedContext = oldStylesContext;
            oldStylesContext = null;
            return new TempSheetStylesBuilderUtils(backingStream, copiedArchive, copiedContext, builder);
        }
        finally
        {
#if SYNC_ONLY
            oldStylesContext?.Dispose();
#else
            var ctxDisposeTask = oldStylesContext?.DisposeAsync();
            if (ctxDisposeTask.HasValue) await ctxDisposeTask.Value.ConfigureAwait(false);
#endif
        }
    }

    [CreateSyncVersion]
    private async Task CopyEntryAsync(ZipArchiveEntry entry, CancellationToken cancellationToken = default)
    {
        var newEntry = _archive.CreateEntry(entry.FullName, CompressionLevel.Fastest);

#if NET8_0_OR_GREATER
        var oldEntryStream = await _oldArchive!.GetEntry(entry.FullName)!.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var oldDisposableSheetStream = oldEntryStream.ConfigureAwait(false);

        var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var newDisposableSheetStream = newEntryStream.ConfigureAwait(false);
#else
        using var oldEntryStream = await _oldArchive!.GetEntry(entry.FullName)!.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var newEntryStream = await newEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#endif

        await oldEntryStream.CopyToAsync(newEntryStream, 81920, cancellationToken).ConfigureAwait(false);
        await newEntryStream.FlushAsync(cancellationToken).ConfigureAwait(false);
    }
    
    [CreateSyncVersion]
    private async Task CopyAndUpdateContentTypesAsync(CancellationToken cancellationToken = default)
    {
        var contentTypesZipEntry = _oldArchive!.Entries.SingleOrDefault(s => s.FullName == ExcelFileNames.ContentTypes);
        if (contentTypesZipEntry is null)
        {
            await GenerateContentTypesXmlAsync(cancellationToken).ConfigureAwait(false);
            return;
        }

#if NET8_0_OR_GREATER
        var stream = await contentTypesZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);
        var doc = await XDocument.LoadAsync(stream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var stream = contentTypesZipEntry.Open();
        var doc = XDocument.Load(stream);
#endif

        var ns = doc.Root!.GetDefaultNamespace();
        var typesElement = doc.Descendants(ns + "Types").Single();

        var partNames = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        var attrNames = typesElement.Elements(ns + "Override").Select(s => s.Attribute("PartName")?.Value);
        foreach (var partName in attrNames.OfType<string>())
        {
            partNames.Add(partName);
        }

        foreach (var (entry, contentType) in _zipContentsMap)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var entryPath = $"/{entry}";
            if (!partNames.Contains(entryPath))
            {
                var newElement = new XElement(ns + "Override", new XAttribute("ContentType", contentType), new XAttribute("PartName", entryPath));
                typesElement.Add(newElement);
            }
        }

        var contentTypesEntry = _archive.CreateEntry(ExcelFileNames.ContentTypes, CompressionLevel.Fastest);
        var contentTypesEntryStream = await contentTypesEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
#if NET8_0_OR_GREATER
        await using var disposableContetTypesEntryStream = contentTypesEntryStream.ConfigureAwait(false);
        await doc.SaveAsync(contentTypesEntryStream, SaveOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var disposableContetTypesEntryStream = contentTypesEntryStream;
        doc.Save(contentTypesEntryStream);
#endif
    }
    
    private class TempSheetStylesBuilderUtils(Stream backingStream, ZipArchive archive, SheetStyleBuilderContext sheetStyleBuilderContext, ISheetStyleBuilder sheetStyleBuilder) : IDisposable, IAsyncDisposable
    {
        private readonly Stream _backingStream = backingStream;

        public ZipArchive Archive { get; } = archive;
        public ISheetStyleBuilder SheetStyleBuilder { get; } = sheetStyleBuilder;
        public SheetStyleBuilderContext SheetStyleBuilderContext { get; } = sheetStyleBuilderContext;

        public void Dispose()
        {
            SheetStyleBuilderContext.Dispose();
            Archive.Dispose();
            _backingStream.Dispose();
        }

        public async ValueTask DisposeAsync()
        {
            await SheetStyleBuilderContext.DisposeAsync().ConfigureAwait(false);
#if NET10_0_OR_GREATER
            await Archive.DisposeAsync().ConfigureAwait(false);
#else
            Archive.Dispose();
#endif
#if NET8_0_OR_GREATER
            await _backingStream.DisposeAsync().ConfigureAwait(false);
#else
            _backingStream.Dispose();
#endif
        }
    }
}
