using System.Drawing;
using MiniExcelLib.OpenXml.Styles;

// ReSharper disable once CheckNamespace
namespace MiniExcelLib.OpenXml;

public sealed class OpenXmlEditor
{
    private const int MaxColumn = 16_384;
    private const int MaxRow = 1_048_576;
    private readonly string? _path;
    private readonly Stream? _stream;
    private readonly List<CellStyleUpdate> _styleUpdates = [];

    internal OpenXmlEditor(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("The path cannot be null or whitespace.", nameof(path));

        _path = path;
    }

    internal OpenXmlEditor(Stream stream)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
    }

    /// <summary>Queues a partial style update for an existing cell.</summary>
    public OpenXmlEditor UpdateCellStyle(string cellReference, Action<OpenXmlCellStyle> update, string? sheetName = null)
    {
        if (update is null)
            throw new ArgumentNullException(nameof(update));

        var style = new OpenXmlCellStyle();
        update(style);
        return UpdateCellStyle(cellReference, style, sheetName);
    }

    /// <summary>Queues a partial style update for an existing cell.</summary>
    public OpenXmlEditor UpdateCellStyle(string cellReference, OpenXmlCellStyle style, string? sheetName = null)
    {
        if (style is null)
            throw new ArgumentNullException(nameof(style));

        if (!CellReferenceConverter.TryParseCellReference(cellReference, out var column, out var row)
            || column > MaxColumn || row > MaxRow)
            throw new ArgumentException($"'{cellReference}' is not a valid cell reference.", nameof(cellReference));

        if (style.FontColor is not { } fontColor)
            throw new ArgumentException("At least one style property must be specified.", nameof(style));

        var normalizedReference = CellReferenceConverter.GetCellFromCoordinates(column, row);
        _styleUpdates.Add(new CellStyleUpdate(normalizedReference, column, row, sheetName, fontColor));
        return this;
    }

    /// <summary>Applies all queued updates to the workbook.</summary>
    public void Save(CancellationToken cancellationToken = default) =>
        SaveAsync(cancellationToken).GetAwaiter().GetResult();

    /// <summary>Applies all queued updates to the workbook asynchronously.</summary>
    public async Task SaveAsync(CancellationToken cancellationToken = default)
    {
        if (_styleUpdates.Count == 0)
            return;

        if (_path is not null)
        {
            await SavePathAsync(cancellationToken).ConfigureAwait(false);
        }
        else
        {
            await SaveStreamAsync(_stream!, cancellationToken).ConfigureAwait(false);
        }

        _styleUpdates.Clear();
    }

    private async Task SavePathAsync(CancellationToken cancellationToken)
    {
        var temporaryPath = $"{_path}.{Guid.NewGuid():N}.tmp";
        try
        {
            using (var source = new FileStream(_path!, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var temporaryStream = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
            {
                await ApplyUpdatesAsync(source, temporaryStream, cancellationToken).ConfigureAwait(false);
                await temporaryStream.FlushAsync(cancellationToken).ConfigureAwait(false);
            }

            ReplaceFile(temporaryPath, _path!);
        }
        finally
        {
            if (File.Exists(temporaryPath))
                File.Delete(temporaryPath);
        }
    }

    private async Task SaveStreamAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (!stream.CanRead || !stream.CanWrite || !stream.CanSeek)
            throw new ArgumentException("The stream must be readable, writable, and seekable.", nameof(stream));

        stream.Seek(0, SeekOrigin.Begin);
        var temporaryPath = Path.GetTempFileName();
        try
        {
            using var temporaryStream = new FileStream(temporaryPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
            await ApplyUpdatesAsync(stream, temporaryStream, cancellationToken).ConfigureAwait(false);

            temporaryStream.Position = 0;
            stream.Position = 0;
            stream.SetLength(0);
            await temporaryStream.CopyToAsync(stream, 81920, cancellationToken).ConfigureAwait(false);
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            File.Delete(temporaryPath);
        }
    }

    private async Task ApplyUpdatesAsync(Stream inputStream, Stream outputStream, CancellationToken cancellationToken)
    {
        inputStream.Seek(0, SeekOrigin.Begin);
        using var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);

        var contentTypes = await LoadDocumentAsync(GetRequiredEntry(inputArchive, ExcelFileNames.ContentTypes), cancellationToken).ConfigureAwait(false);
        if (contentTypes.Descendants().Attributes("ContentType")
            .Any(attribute => attribute.Value.IndexOf("macroEnabled", StringComparison.OrdinalIgnoreCase) >= 0))
            throw new NotSupportedException("MiniExcel's OpenXml editor does not support the .xlsm format.");

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

        using var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var inputEntry in inputArchive.Entries)
        {
            if (inputEntry.FullName.Equals(ExcelFileNames.Styles, StringComparison.OrdinalIgnoreCase)
                || worksheetPaths.Contains(inputEntry.FullName))
                continue;

            await CopyEntryAsync(inputEntry, outputArchive, cancellationToken).ConfigureAwait(false);
        }

        foreach (var sheetUpdates in updatesBySheet)
        {
            var inputEntry = GetRequiredEntry(inputArchive, sheetUpdates.Key);
            var outputEntry = CreateEntry(outputArchive, inputEntry);
            await RewriteWorksheetAsync(inputEntry, outputEntry, sheetUpdates, styleContext, cancellationToken).ConfigureAwait(false);
        }

        await WriteDocumentEntryAsync(outputArchive, stylesEntry, styles, cancellationToken).ConfigureAwait(false);
    }

    private static void ReplaceFile(string sourcePath, string destinationPath)
    {
        File.Replace(sourcePath, destinationPath, null);
    }

    private List<ResolvedStyleUpdate> ResolveUpdates(IReadOnlyList<SheetReference> sheets)
    {
        var updates = new Dictionary<(string SheetPath, string CellReference), ResolvedStyleUpdate>();

        foreach (var update in _styleUpdates)
        {
            var sheet = update.SheetName is null
                ? sheets.FirstOrDefault()
                : sheets.FirstOrDefault(candidate => string.Equals(candidate.Name, update.SheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
                throw new ArgumentException(update.SheetName is null
                    ? "The workbook does not contain any worksheets."
                    : $"Worksheet '{update.SheetName}' does not exist.");

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
            .Where(element => element.Name.LocalName == "Relationship")
            .Where(element => element.Attribute("Type")?.Value.EndsWith("/worksheet", StringComparison.Ordinal) == true)
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

    private static ZipArchiveEntry GetRequiredEntry(ZipArchive archive, string path) =>
        archive.GetEntry(path) ?? throw new InvalidDataException($"The OpenXml document does not contain '{path}'.");

    private static async Task<XDocument> LoadDocumentAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
    {
        using var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        return await XDocument.LoadAsync(stream, LoadOptions.PreserveWhitespace, cancellationToken).ConfigureAwait(false);
    }

    private static async Task CopyEntryAsync(ZipArchiveEntry inputEntry, ZipArchive outputArchive, CancellationToken cancellationToken)
    {
        var outputEntry = CreateEntry(outputArchive, inputEntry);
        if (inputEntry.FullName.EndsWith("/", StringComparison.Ordinal))
            return;

        using var inputStream = await inputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await inputStream.CopyToAsync(outputStream, 81920, cancellationToken).ConfigureAwait(false);
    }

    private static async Task RewriteWorksheetAsync(ZipArchiveEntry inputEntry, ZipArchiveEntry outputEntry,
        IEnumerable<ResolvedStyleUpdate> updates, StyleUpdateContext styleContext, CancellationToken cancellationToken)
    {
        var pendingUpdates = updates.ToDictionary(update => update.CellReference, StringComparer.OrdinalIgnoreCase);
        using var inputStream = await inputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        using var reader = XmlReader.Create(inputStream, new XmlReaderSettings
        {
            Async = true,
            CloseInput = false,
            DtdProcessing = DtdProcessing.Prohibit
        });
        using var writer = XmlWriter.Create(outputStream, new XmlWriterSettings
        {
            Async = true,
            CloseOutput = false,
            Encoding = new UTF8Encoding(false),
            Indent = false
        });

        while (await reader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c"
                && reader.GetAttribute("r") is { } cellReference
                && pendingUpdates.TryGetValue(cellReference, out var update))
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

    private static async Task WriteCellStartElementAsync(XmlReader reader, XmlWriter writer, int styleIndex)
    {
        await writer.WriteStartElementAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI).ConfigureAwait(false);
        var wroteStyle = false;
        if (reader.MoveToFirstAttribute())
        {
            do
            {
                if (reader.LocalName == "s" && reader.NamespaceURI.Length == 0)
                {
                    await writer.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);
                    wroteStyle = true;
                }
                else
                {
                    await writer.WriteAttributeStringAsync(reader.Prefix, reader.LocalName, reader.NamespaceURI, reader.Value).ConfigureAwait(false);
                }
            }
            while (reader.MoveToNextAttribute());

            reader.MoveToElement();
        }

        if (!wroteStyle)
            await writer.WriteAttributeStringAsync(null, "s", null, styleIndex.ToString(CultureInfo.InvariantCulture)).ConfigureAwait(false);

        if (reader.IsEmptyElement)
            await writer.WriteEndElementAsync().ConfigureAwait(false);
    }

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

    private static async Task WriteDocumentEntryAsync(ZipArchive outputArchive, ZipArchiveEntry inputEntry,
        XDocument document, CancellationToken cancellationToken)
    {
        var outputEntry = CreateEntry(outputArchive, inputEntry);
        using var outputStream = await outputEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
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
            var color = new XElement(_namespace + "color",
                new XAttribute("rgb", $"{fontColor.A:X2}{fontColor.R:X2}{fontColor.G:X2}{fontColor.B:X2}"));
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