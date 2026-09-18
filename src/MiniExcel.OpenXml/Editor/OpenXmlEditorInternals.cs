using System.Drawing;
using MiniExcelLib.OpenXml.Styles;

namespace MiniExcelLib.OpenXml.Editor;

public sealed partial class OpenXmlEditorInternals
{
    private const int MaxColumn = 16_384;
    private const int MaxRow = 1_048_576;

    private readonly Stream _stream;
    private readonly bool _leaveOpen;
    private readonly List<CellStyleUpdate> _styleUpdates = [];

    internal OpenXmlEditorInternals(Stream stream, bool leaveOpen)
    {
        if (!stream.CanRead || !stream.CanWrite || !stream.CanSeek)
            throw new ArgumentException("The stream must be readable, writable, and seekable.", nameof(stream));

        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
    }

    /// <summary>Queues a partial style update for an existing cell.</summary>
    internal void UpdateCellStyle(string cellReference, OpenXmlCellStyle cellStyle, string? sheetName = null)
    {
        if (cellStyle is null)
            throw new ArgumentNullException(nameof(cellStyle));

        if (!CellReferenceConverter.TryParseCellReference(cellReference, out var column, out var row) || column > MaxColumn || row > MaxRow)
            throw new ArgumentException($"'{cellReference}' is not a valid cell reference.", nameof(cellReference));

        if (cellStyle.FontColor is not { } fontColor)
            throw new ArgumentException("At least one style property must be specified.", nameof(cellStyle));

        var normalizedReference = CellReferenceConverter.GetCellFromCoordinates(column, row);
        _styleUpdates.Add(new CellStyleUpdate(normalizedReference, column, row, sheetName, fontColor));
    }

    /// <summary>Applies all queued updates to the workbook and saves it to an output stream.</summary>
    [CreateSyncVersion]
    public async Task SaveAsync(Stream outputStream, CancellationToken cancellationToken = default)
    {
        try
        {
            if (_styleUpdates.Count == 0)
                return;

            _stream.Seek(0, SeekOrigin.Begin);
            await ApplyUpdatesAsync(_stream, outputStream, cancellationToken).ConfigureAwait(false);
            await outputStream.FlushAsync(cancellationToken).ConfigureAwait(false);

            _styleUpdates.Clear();
        }
        finally
        {
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }

    /// <summary>Applies all queued updates to the workbook and replaces the original stream.</summary>
    [CreateSyncVersion]
    public async Task SaveAsync(CancellationToken cancellationToken = default)
    {
        try
        {
            if (_styleUpdates.Count == 0)
                return;

            _stream.Seek(0, SeekOrigin.Begin);

            var tempStream = new MemoryStream();
            await using var disposableMemoryStream = tempStream.ConfigureAwait(false);

            await ApplyUpdatesAsync(_stream, tempStream, cancellationToken).ConfigureAwait(false);

            cancellationToken.ThrowIfCancellationRequested();
            // We cannot honor the cancellation of the task after this point because
            // the workbook would be only partially written to the stream thus corrupting the original document

            _stream.Seek(0, SeekOrigin.Begin);
            _stream.SetLength(0);
            
            tempStream.Seek(0, SeekOrigin.Begin);
            await tempStream.CopyToAsync(_stream, 81920, CancellationToken.None).ConfigureAwait(false);
            await _stream.FlushAsync(CancellationToken.None).ConfigureAwait(false);

            _styleUpdates.Clear();
        }
        finally
        {
            if (!_leaveOpen)
            {
                await _stream.DisposeAsync().ConfigureAwait(false);
            }
        }
    }

    [CreateSyncVersion]
    private async Task ApplyUpdatesAsync(Stream inputStream, Stream outputStream, CancellationToken cancellationToken)
    {
        inputStream.Seek(0, SeekOrigin.Begin);

#if NET10_0_OR_GREATER
        var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);
        await using var disposableInputArchive = inputArchive.ConfigureAwait(false);
#else
        using var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);
#endif

        var contentTypes = await LoadDocumentAsync(GetRequiredEntry(inputArchive, ExcelFileNames.ContentTypes), cancellationToken).ConfigureAwait(false);
        if (contentTypes.Descendants()
            .Attributes("ContentType")
            .Any(attribute => attribute.Value.Contains("macroEnabled", StringComparison.OrdinalIgnoreCase)))
        {
            throw new NotSupportedException("MiniExcel's OpenXmlEditor does not support the .xlsm format.");
        }

        var workbook = await LoadDocumentAsync(GetRequiredEntry(inputArchive, ExcelFileNames.Workbook), cancellationToken).ConfigureAwait(false);
        var workbookRelationships = await LoadDocumentAsync(GetRequiredEntry(inputArchive, ExcelFileNames.WorkbookRels), cancellationToken).ConfigureAwait(false);
        var sheets = GetSheets(workbook, workbookRelationships);
        var pendingUpdates = ResolveUpdates(sheets);

        var stylesEntry = GetRequiredEntry(inputArchive, ExcelFileNames.Styles);
        var styles = await LoadDocumentAsync(stylesEntry, cancellationToken).ConfigureAwait(false);
        var styleContext = new StyleUpdateContext(styles);
        var updatesBySheet = pendingUpdates
            .GroupBy(update => update.Sheet.Path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var worksheetPaths = new HashSet<string>(
            updatesBySheet.Select(group => group.Key),
            StringComparer.OrdinalIgnoreCase);

#if NET10_0_OR_GREATER
        var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
        await using var disposableOutputArchive = outputArchive.ConfigureAwait(false);
#else
        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
#endif
        
        foreach (var inputEntry in inputArchive.Entries)
        {
            if (!inputEntry.FullName.Equals(ExcelFileNames.Styles, StringComparison.OrdinalIgnoreCase) &&
                !worksheetPaths.Contains(inputEntry.FullName))
            {
                await CopyEntryAsync(inputEntry, outputArchive, cancellationToken).ConfigureAwait(false);
            }
        }

        foreach (var sheetUpdates in updatesBySheet)
        {
            var inputEntry = GetRequiredEntry(inputArchive, sheetUpdates.Key);
            var outputEntry = CreateEntry(outputArchive, inputEntry);
            await RewriteWorksheetAsync(inputEntry, outputEntry, sheetUpdates, styleContext, cancellationToken).ConfigureAwait(false);
        }

        await WriteDocumentEntryAsync(outputArchive, stylesEntry, styles, cancellationToken).ConfigureAwait(false);
    }

    private List<ResolvedStyleUpdate> ResolveUpdates(IReadOnlyList<SheetReference> sheets)
    {
        var updates = new Dictionary<(string SheetPath, string CellReference), ResolvedStyleUpdate>();

        foreach (var update in _styleUpdates)
        {
            var sheet = string.IsNullOrEmpty(update.SheetName)
                ? sheets.FirstOrDefault()
                : sheets.FirstOrDefault(sheet => string.Equals(sheet.Name, update.SheetName, StringComparison.OrdinalIgnoreCase));

            if (sheet is null)
            {
                var errorMsg = update.SheetName is null
                    ? "The workbook does not contain any worksheets."
                    : $"Worksheet '{update.SheetName}' does not exist.";
                
                throw new ArgumentException(errorMsg);
            }

            updates[(sheet.Path, update.CellReference)] = new ResolvedStyleUpdate(
                sheet, update.CellReference, update.Column, update.Row, update.FontColor);
        }

        return updates.Values
            .OrderBy(update => update.Sheet.Index)
            .ThenBy(update => update.Row)
            .ThenBy(update => update.Column)
            .ToList();
    }

    private static List<SheetReference> GetSheets(XDocument workbook, XDocument relationships)
    {
        var relationshipTargets = relationships.Descendants()
            .Where(element => element.Name.LocalName == "Relationship" && 
                              element.Attribute("Type")?.Value.EndsWith("/worksheet", StringComparison.Ordinal) == true)
            .ToDictionary(
                element => element.Attribute("Id")?.Value ?? string.Empty,
                element => NormalizeWorkbookTarget(element.Attribute("Target")?.Value ?? string.Empty),
                StringComparer.Ordinal);

        return workbook.Descendants()
            .Where(element => element.Name.LocalName == "sheet")
            .Select((element, index) =>
            {
                var relationshipId = element.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "id")?.Value
                    ?? throw new InvalidDataException("A worksheet is missing its relationship id.");

                if (!relationshipTargets.TryGetValue(relationshipId, out var path))
                    throw new InvalidDataException($"Worksheet relationship '{relationshipId}' does not exist.");

                return new SheetReference(index, element.Attribute("name")?.Value ?? string.Empty, path);
            })
            .ToList();
    }

    private static string NormalizeWorkbookTarget(string target)
    {
        if (string.IsNullOrWhiteSpace(target))
            throw new InvalidDataException("A worksheet relationship has an empty target.");

        var uri = new Uri(new Uri("https://miniexcel.local/xl/workbook.xml"), target.Replace('\\', '/'));
        return Uri.UnescapeDataString(uri.AbsolutePath).TrimStart('/');
    }

    private static int ParseStyleIndex(string? value, string cellReference)
    {
        if (value is null)
            return 0;

        if (int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var styleIndex) && styleIndex >= 0)
            return styleIndex;

        throw new InvalidDataException($"Cell '{cellReference}' has an invalid style index.");
    }

    private static ZipArchiveEntry GetRequiredEntry(ZipArchive archive, string path) 
        => archive.GetEntry(path) ?? throw new InvalidDataException($"The OpenXml document does not contain '{path}'.");

    [CreateSyncVersion]
    private static async Task<XDocument> LoadDocumentAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);
        return await XDocument.LoadAsync(stream, LoadOptions.PreserveWhitespace, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task CopyEntryAsync(ZipArchiveEntry inputEntry, ZipArchive outputArchive, CancellationToken cancellationToken)
    {
        var outputEntry = CreateEntry(outputArchive, inputEntry);
        if (inputEntry.FullName.EndsWith("/", StringComparison.Ordinal))
            return;

        var inputStream = await inputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableInputStream = inputStream.ConfigureAwait(false);
        
        var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableOutputStream = outputStream.ConfigureAwait(false);

        await inputStream.CopyToAsync(outputStream, 81920, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task RewriteWorksheetAsync(ZipArchiveEntry inputEntry, ZipArchiveEntry outputEntry,
        IEnumerable<ResolvedStyleUpdate> updates, StyleUpdateContext styleContext, CancellationToken cancellationToken)
    {
        var pendingUpdates = updates.ToDictionary(update => update.CellReference, StringComparer.OrdinalIgnoreCase);

        var inputStream = await inputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableInputStream = inputStream.ConfigureAwait(false);
        
        var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableOutputStream = outputStream.ConfigureAwait(false);

        var readerSettings = new XmlReaderSettings
        {
#if !SYNC_ONLY            
            Async = true,
#endif
            XmlResolver = null
        };
        using var reader = XmlReader.Create(inputStream, readerSettings);

        var writerSettings = new XmlWriterSettings
        {
#if !SYNC_ONLY            
            Async = true,
#endif
            Encoding = new UTF8Encoding(false)
        };
#if NET
        var writer = XmlWriter.Create(outputStream, writerSettings);
        await using var disposableWriter = writer.ConfigureAwait(false);
#else
        using var writer = XmlWriter.Create(outputStream, writerSettings);
#endif

        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader is { NodeType: XmlNodeType.Element, LocalName: "c" } && 
                reader.GetAttribute("r") is { } cellReference &&
                pendingUpdates.TryGetValue(cellReference, out var update))
            {
                var originalStyleIndex = ParseStyleIndex(reader.GetAttribute("s"), cellReference);
                var styleIndex = styleContext.GetStyleIndex(originalStyleIndex, update.FontColor);
                await WriteCellStartElementAsync(reader, writer, styleIndex).ConfigureAwait(false);
                pendingUpdates.Remove(cellReference);

                continue;
            }

            await WriteCurrentNodeAsync(reader, writer).ConfigureAwait(false);
        }

        if (pendingUpdates.Count > 0)
        {
            var missingCell = pendingUpdates.Values.OrderBy(update => update.Row).ThenBy(update => update.Column).First();
            throw new InvalidDataException($"Cell '{missingCell.CellReference}' does not exist in worksheet '{missingCell.Sheet.Name}'.");
        }

        await writer.FlushAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task WriteCellStartElementAsync(XmlReader reader, XmlWriter writer, int styleIndex)
    {
        await writer.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI).ConfigureAwait(false);
        var styleWritten = false;

        if (reader.MoveToFirstAttribute())
        {
            do
            {
                if (reader is { LocalName: "s", NamespaceURI.Length: 0 })
                {
                    await writer.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);
                    styleWritten = true;
                }
                else
                {
                    await writer.WriteAttributeStringAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value).ConfigureAwait(false);
                }
            }
            while (reader.MoveToNextAttribute());

            reader.MoveToElement();
        }

        if (!styleWritten)
            await writer.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);

        if (reader.IsEmptyElement)
            await writer.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task WriteCurrentNodeAsync(XmlReader reader, XmlWriter writer)
    {
        switch (reader.NodeType)
        {
            case XmlNodeType.Element:
                await writer.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI).ConfigureAwait(false);
                if (reader.MoveToFirstAttribute())
                {
                    do
                    {
                        await writer.WriteAttributeStringAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value).ConfigureAwait(false);
                    }
                    while (reader.MoveToNextAttribute());

                    reader.MoveToElement();
                }
                if (reader.IsEmptyElement)
                    await writer.WriteEndElementAsync().ConfigureAwait(false);
                break;

            case XmlNodeType.EndElement:
                await writer.WriteFullEndElementAsync().ConfigureAwait(false);
                break;

            case XmlNodeType.Text:
                await writer.WriteStringAsync(reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.CDATA:
                await writer.WriteCDataAsync(reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.Whitespace:
            case XmlNodeType.SignificantWhitespace:
                await writer.WriteWhitespaceAsync(reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.Comment:
                await writer.WriteCommentAsync(reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.ProcessingInstruction:
                await writer.WriteProcessingInstructionAsync(reader.Name, reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.XmlDeclaration:
                await writer.WriteStartDocumentAsync().ConfigureAwait(false);
                break;

            case XmlNodeType.DocumentType:
                await writer.WriteDocTypeAsync(reader.Name, reader.GetAttribute("PUBLIC"), reader.GetAttribute("SYSTEM"), reader.Value).ConfigureAwait(false);
                break;

            case XmlNodeType.EntityReference:
                await writer.WriteEntityRefAsync(reader.Name).ConfigureAwait(false);
                break;
        }
    }

    [CreateSyncVersion]
    private static async Task WriteDocumentEntryAsync(ZipArchive outputArchive, ZipArchiveEntry inputEntry,
        XDocument document, CancellationToken cancellationToken)
    {
        var outputEntry = CreateEntry(outputArchive, inputEntry);
        var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableOuputStream = outputStream.ConfigureAwait(false);

        await document.SaveAsync(outputStream, SaveOptions.DisableFormatting, cancellationToken).ConfigureAwait(false);
    }

    private static ZipArchiveEntry CreateEntry(ZipArchive archive, ZipArchiveEntry sourceEntry)
    {
        var entry = archive.CreateEntry(sourceEntry.FullName, CompressionLevel.Optimal);
        entry.LastWriteTime = sourceEntry.LastWriteTime;
        return entry;
    }

    private sealed class StyleUpdateContext
    {
        private readonly XNamespace _namespace;
        private readonly XElement _fonts;
        private readonly XElement _cellFormats;
        private readonly List<XElement> _originalFonts;
        private readonly List<XElement> _originalCellFormats;
        private readonly Dictionary<(int StyleIndex, int Argb), int> _styleIndexes = [];

        internal StyleUpdateContext(XDocument styles)
        {
            var root = styles.Root ?? throw new InvalidDataException("The styles document has no root element.");
            _namespace = root.Name.Namespace;
            _fonts = root.Element(_namespace + "fonts") ?? throw new InvalidDataException("The styles document has no fonts collection.");
            _cellFormats = root.Element(_namespace + "cellXfs") ?? throw new InvalidDataException("The styles document has no cell formats collection.");
            _originalFonts = _fonts.Elements(_namespace + "font").ToList();
            _originalCellFormats = _cellFormats.Elements(_namespace + "xf").ToList();
        }

        internal int GetStyleIndex(int originalStyleIndex, Color fontColor)
        {
            var key = (originalStyleIndex, fontColor.ToArgb());
            if (_styleIndexes.TryGetValue(key, out var styleIndex))
                return styleIndex;

            if (originalStyleIndex >= _originalCellFormats.Count)
                throw new InvalidDataException($"Style index '{originalStyleIndex}' does not exist.");

            var originalCellFormat = _originalCellFormats[originalStyleIndex];
            var fontIdValue = originalCellFormat.Attribute("fontId")?.Value ?? "0";
            if (!int.TryParse(fontIdValue, NumberStyles.None, CultureInfo.InvariantCulture, out var fontId) || fontId < 0 || fontId >= _originalFonts.Count)
                throw new InvalidDataException($"Font index '{fontIdValue}' does not exist.");

            var font = new XElement(_originalFonts[fontId]);
            var color = new XElement(_namespace + "color", new XAttribute("rgb", $"{fontColor.A:X2}{fontColor.R:X2}{fontColor.G:X2}{fontColor.B:X2}"));
            var oldColor = font.Elements().FirstOrDefault(element => element.Name.LocalName == "color");

            if (oldColor is null)
                font.Add(color);
            else
                oldColor.ReplaceWith(color);

            _fonts.Add(font);
            _fonts.SetAttributeValue("count", _fonts.Elements(_namespace + "font").Count());
            var newFontId = _fonts.Elements(_namespace + "font").Count() - 1;

            var cellFormat = new XElement(originalCellFormat);
            cellFormat.SetAttributeValue("fontId", newFontId);
            cellFormat.SetAttributeValue("applyFont", "1");
            _cellFormats.Add(cellFormat);
            _cellFormats.SetAttributeValue("count", _cellFormats.Elements(_namespace + "xf").Count());
            
            styleIndex = _cellFormats.Elements(_namespace + "xf").Count() - 1;
            _styleIndexes.Add(key, styleIndex);

            return styleIndex;
        }
    }

    [CreateSyncVersion]
    /* Todo: this method is not very efficient, but workbook.xml is generally a very small file so at the moment it's not worth over-optimizing it.
     Also, consider adding active sheet as one of the editable properties.*/
    internal async Task AlterWorksheetAsync(string sheetName, string? newSheetName, int? newSheetIndex, SheetState? newSheetState, CancellationToken cancellationToken = default)
    {
        if (newSheetName is null && newSheetIndex is null && newSheetState is null)
            return;

        var archive = await ZipArchive.CreateAsync(_stream, ZipArchiveMode.Update, true, new UTF8Encoding(true), cancellationToken).ConfigureAwait(false);
        var oldWorkbookEntry = archive.GetEntry(ExcelFileNames.Workbook)!;
 
        try
        {
            var xmlDoc = await LoadWorkbook().ConfigureAwait(false);

            oldWorkbookEntry.Delete();
            // We cannot honor the cancellation of the task after this point because the worbook would get corrputed
            var newWorkbookEntry = archive.CreateEntry(ExcelFileNames.Workbook, CompressionLevel.Fastest);
            
            var newZipStream = await newWorkbookEntry.OpenAsync(CancellationToken.None).ConfigureAwait(false);
            await using var newDisposableZipStream = newZipStream.ConfigureAwait(false);
#if NET
            var writer = XmlWriter.Create(newZipStream, new XmlWriterSettings
            {
#if !SYNC_ONLY
                Async = true
#endif
            });
            await using var disposableWriter = writer.ConfigureAwait(false);
            await xmlDoc.WriteToAsync(writer, CancellationToken.None).ConfigureAwait(false);
#else
            using var writer = XmlWriter.Create(newZipStream, new XmlWriterSettings { Async = false });
            xmlDoc.WriteTo(writer);
#endif
        }
        finally
        {
#if NET10_0_OR_GREATER
            await archive.DisposeAsync().ConfigureAwait(false);
#else
            archive.Dispose();
#endif
        }
        return;

        async Task<XDocument> LoadWorkbook()
        {
            var zipStream = await oldWorkbookEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
            await using var disposableZipStream = zipStream.ConfigureAwait(false);

            var workbookDoc = await XDocument.LoadAsync(zipStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
            var sheetsContainer = workbookDoc.Root?.Element((XNamespace)Schemas.SpreadsheetmlXmlMain + "sheets")!;
            var sheets = sheetsContainer.Elements().ToList();

            if (sheets.Find(s => s.Attribute("name")?.Value.Equals(sheetName, StringComparison.OrdinalIgnoreCase) is true) is not { } sheet)
                throw new InvalidDataException($"Sheet {sheetName} not found");

            if (newSheetName is not null)
            {
                ThrowHelper.ThrowIfInvalidSheetName(newSheetName);
                sheet.SetAttributeValue("name", newSheetName);
            }

            if (newSheetIndex is not null)
            {
                var newIndex = Math.Clamp(newSheetIndex.Value, 0, sheets.Count - 1);
                sheets.Remove(sheet);
                sheets.Insert(newIndex, sheet);

                sheetsContainer.RemoveAll();
                sheetsContainer.Add(sheets);
            }

            if (newSheetState is not null)
            {
                sheet.SetAttributeValue("state", newSheetState switch
                {
                    SheetState.Visible => "visible",
                    SheetState.Hidden => "hidden",
                    SheetState.VeryHidden => "veryHidden",
                    _ => "visible"
                });
            }

            return workbookDoc;
        }
    }
    
    private sealed class CellStyleUpdate(string cellReference, int column, int row, string? sheetName, Color fontColor)
    {
        internal string CellReference { get; } = cellReference;
        internal int Column { get; } = column;
        internal int Row { get; } = row;
        internal string? SheetName { get; } = sheetName;
        internal Color FontColor { get; } = fontColor;
    }

    private sealed class SheetReference(int index, string name, string path)
    {
        internal int Index { get; } = index;
        internal string Name { get; } = name;
        internal string Path { get; } = path;
    }

    private sealed class ResolvedStyleUpdate(SheetReference sheet, string cellReference, int column, int row, Color fontColor)
    {
        internal SheetReference Sheet { get; } = sheet;
        internal string CellReference { get; } = cellReference;
        internal int Column { get; } = column;
        internal int Row { get; } = row;
        internal Color FontColor { get; } = fontColor;
    }
}
