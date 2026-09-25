namespace MiniExcelLib.OpenXml.Templates;

/// <summary>
/// Template image support: byte[] values resolved from template placeholders are emitted as embedded
/// images, reusing the same OpenXML primitives as the SaveAs pipeline (ImageHelper format detection,
/// FileDto and the ExcelXml drawing builders). See issue #972 / #604.
/// </summary>
internal partial class OpenXmlTemplate
{
    internal const string ImageMarkerPrefix = "@@@imageid@@@,";

    private readonly List<FileDto> _files = [];
    private readonly Dictionary<string, PendingImage> _pendingImages = [];
    private readonly Dictionary<int, (string DrawingPath, string DrawingRelsPath)> _sheetTemplateDrawings = [];
    private readonly Dictionary<int, string> _sheetTemplateRels = [];
    private readonly List<string> _createdDrawingParts = [];
    private int _currentSheetIndex;
    private int _nextImageId;

    private sealed class PendingImage(byte[] bytes, string extension, (int Width, int Height)? size)
    {
        internal byte[] Bytes { get; } = bytes;
        internal string Extension { get; } = extension;
        internal (int Width, int Height)? Size { get; } = size;
    }

#if NET
    [GeneratedRegex(@"<[A-Za-z0-9:]*c\b[^>]*\br=""(?<ref>[A-Z]+[0-9]+)""[^>]*>(?:(?!</[A-Za-z0-9:]*c>).)*?@@@imageid@@@,(?<id>[0-9]+)(?:(?!</[A-Za-z0-9:]*c>).)*?</[A-Za-z0-9:]*c>", RegexOptions.Singleline)]
    private static partial Regex ImageMarkerCellRegex();

    [GeneratedRegex(@"^\s*<[A-Za-z0-9:]*row\b[^>]*\sht=""(?<ht>[0-9]+(?:\.[0-9]+)?)""")]
    private static partial Regex ImageRowHeightRegex();

    private static readonly Regex ImageMarkerCellRegexImpl = ImageMarkerCellRegex();
    private static readonly Regex ImageRowHeightRegexImpl = ImageRowHeightRegex();
#else
    private static readonly Regex ImageMarkerCellRegexImpl = new(
        @"<[A-Za-z0-9:]*c\b[^>]*\br=""(?<ref>[A-Z]+[0-9]+)""[^>]*>(?:(?!</[A-Za-z0-9:]*c>).)*?@@@imageid@@@,(?<id>[0-9]+)(?:(?!</[A-Za-z0-9:]*c>).)*?</[A-Za-z0-9:]*c>",
        RegexOptions.Compiled | RegexOptions.Singleline);

    private static readonly Regex ImageRowHeightRegexImpl = new(
        @"^\s*<[A-Za-z0-9:]*row\b[^>]*\sht=""(?<ht>[0-9]+(?:\.[0-9]+)?)""",
        RegexOptions.Compiled);
#endif

    private const long EmuPerPoint = 12700;

    private void ResetImageState()
    {
        _files.Clear();
        _pendingImages.Clear();
        _sheetTemplateDrawings.Clear();
        _sheetTemplateRels.Clear();
        _createdDrawingParts.Clear();
        _nextImageId = 0;
        _currentSheetIndex = 0;
    }

    /// <summary>
    /// Returns an inline marker for <see cref="byte"/> array values that are recognised images and
    /// registers their bytes for later emission. Values that are not recognised images fall back to
    /// the regular scalar formatting, mirroring <c>SaveAs</c>' byte[] handling as closely as possible.
    /// </summary>
    private string? GetImageMarker(object? value)
    {
        if (value is not byte[] bytes || !_configuration.EnableConvertByteArray)
            return null;

        var format = ImageHelper.GetImageFormat(bytes);
        if (format == ImageHelper.ImageFormat.Unknown)
            return null;

        var id = _nextImageId.ToString(CultureInfo.InvariantCulture);
        _nextImageId++;
        _pendingImages[id] = new PendingImage(bytes, format.ToString().ToLowerInvariant(), ImageHelper.GetImageSize(bytes));
        return ImageMarkerPrefix + id;
    }

    private string GetFormattedValueWithImages(PropertyInfo? propInfo, object? cellValue, Type? type)
        => GetImageMarker(cellValue) ?? GetFormattedValue(propInfo, cellValue, type);

    /// <summary>
    /// Walks the remaining segments of a dotted placeholder expression from an already resolved root
    /// value, so that nested scalars (e.g. {{Company.Logo}}) can be resolved. Index 0 is the root
    /// property, which is the <paramref name="root"/> value itself.
    /// </summary>
    private static bool TryResolvePropertyPath(object? root, string[] segments, out object? value)
    {
        value = root;
        for (var i = 1; i < segments.Length; i++)
        {
            if (value is null)
                return false;

            var type = value.GetType();
            var property = type.GetProperty(segments[i], BindingFlags.Public | BindingFlags.Instance);
            if (property is not null && property.CanRead && property.GetIndexParameters().Length == 0)
            {
                value = property.GetValue(value);
                continue;
            }

            var field = type.GetField(segments[i], BindingFlags.Public | BindingFlags.Instance);
            if (field is not null)
            {
                value = field.GetValue(value);
                continue;
            }

            value = null;
            return false;
        }

        return true;
    }

    private bool HasImagesForSheet(int sheetIndex)
        => _files.Exists(file => file.SheetIndex == sheetIndex && file.IsImage);

    /// <summary>
    /// Replaces image markers embedded in a rendered row with empty cells, registering each image
    /// against its final cell reference. Coordinates are taken from the cell reference itself, which
    /// the surrounding code has already rewritten to the final (post collection expansion) row.
    /// </summary>
    private string CaptureAndClearImageMarkers(string rowXml, int sheetIndex)
    {
        if (_pendingImages.Count == 0 || !rowXml.Contains(ImageMarkerPrefix))
            return rowXml;

        var rowHeightPoints = GetRowHeightPoints(rowXml);

        while (true)
        {
            var match = ImageMarkerCellRegexImpl.Match(rowXml);
            if (!match.Success)
                return rowXml;

            var id = match.Groups["id"].Value;
            if (CellReferenceConverter.TryParseCellReference(match.Groups["ref"].Value, out var column, out var row) &&
                _pendingImages.TryGetValue(id, out var pending))
            {
                var file = new FileDto
                {
                    SheetIndex = sheetIndex,
                    RowIndex = row,
                    CellIndex = column,
                    Contents = pending.Bytes,
                    Extension = pending.Extension,
                    IsImage = true
                };

                ApplyRowHeightSize(file, pending, rowHeightPoints);
                _files.Add(file);
            }

            var clearedCell = match.Value.Replace(ImageMarkerPrefix + id, string.Empty);
            rowXml = rowXml.Remove(match.Index, match.Length).Insert(match.Index, clearedCell);
        }
    }

    /// <summary>
    /// Sizes an image to the height of the row it is anchored to, preserving its aspect ratio. Rows
    /// without an explicit height keep the default anchor size.
    /// </summary>
    private static void ApplyRowHeightSize(FileDto file, PendingImage pending, double rowHeightPoints)
    {
        if (pending.Size is not { } size || size.Width <= 0 || size.Height <= 0 || rowHeightPoints <= 0)
            return;

        var heightEmu = (long)Math.Round(rowHeightPoints * EmuPerPoint);
        var widthEmu = (long)Math.Round(heightEmu * (size.Width / (double)size.Height));
        file.ImageWidthEmu = widthEmu;
        file.ImageHeightEmu = heightEmu;
    }

    private static double GetRowHeightPoints(string rowXml)
    {
        var match = ImageRowHeightRegexImpl.Match(rowXml);
        return match.Success &&
               double.TryParse(match.Groups["ht"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var height)
            ? height
            : 0;
    }

    private static bool IsDrawingPrecedingElement(XElement element)
        => element.Name.LocalName is "tableParts" or "oleObjects" or "controls" or "extLst";

    [CreateSyncVersion]
    private static async Task WriteDrawingReferenceAsync(XmlWriter writer, string? prefix, int sheetIndex)
    {
        var prefixSeparator = string.IsNullOrEmpty(prefix) ? string.Empty : prefix + ":";
        await writer.WriteRawAsync($"<{prefixSeparator}drawing r:id=\"rDrawing{sheetIndex}\" />").ConfigureAwait(false);
    }

    /// <summary>
    /// Resolves, for every non-parametrized template worksheet that already contains a drawing, the
    /// drawing and drawing-relationships paths so they can be merged with the generated images.
    /// </summary>
    [CreateSyncVersion]
    private static async Task<Dictionary<string, (string DrawingPath, string DrawingRelsPath)>> GetTemplateDrawingPathsAsync(
        ZipArchive templateArchive, IDictionary<string, string> sheetNamesMap, CancellationToken cancellationToken)
    {
        var result = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        var packageRelNs = (XNamespace)Schemas.OpenXmlPackageRelationships;
        var sheetNs = (XNamespace)Schemas.SpreadsheetmlXmlMain;
        var relNs = (XNamespace)Schemas.SpreadsheetmlXmlRelationships;

        foreach (var sheetPath in sheetNamesMap.Keys)
        {
            if (ParametrizedSheetRegexImpl.IsMatch(sheetNamesMap[sheetPath]))
                continue;

            if (templateArchive.GetEntry(sheetPath) is null)
                continue;

            var sheetDoc = await LoadXmlAsync(templateArchive, sheetPath, cancellationToken).ConfigureAwait(false);
            var rId = sheetDoc.Descendants(sheetNs + "drawing").FirstOrDefault()?.Attribute(relNs + "id")?.Value;
            if (string.IsNullOrEmpty(rId))
                continue;

            var relsPath = $"xl/worksheets/_rels/{Path.GetFileName(sheetPath)}.rels";
            if (templateArchive.GetEntry(relsPath) is null)
                continue;

            var relsDoc = await LoadXmlAsync(templateArchive, relsPath, cancellationToken).ConfigureAwait(false);
            var target = relsDoc.Descendants(packageRelNs + "Relationship")
                .FirstOrDefault(rel => rel.Attribute("Id")?.Value == rId)
                ?.Attribute("Target")?.Value;

            if (string.IsNullOrEmpty(target))
                continue;

            var normalized = target!.Replace('\\', '/');
            var drawingPath = normalized.StartsWith("../", StringComparison.Ordinal)
                ? "xl/" + normalized[3..]
                : normalized.TrimStart('/');

            result[sheetPath] = (drawingPath, $"xl/drawings/_rels/{Path.GetFileName(drawingPath)}.rels");
        }

        return result;
    }

    /// <summary>
    /// Emits the image parts (media, drawing, drawing relationships and worksheet relationship) for
    /// every sheet that produced images. Sheets whose template already declared a drawing reuse and
    /// extend that drawing instead of creating a second, unreferenced one.
    /// </summary>
    [CreateSyncVersion]
    private async Task EmitTemplateImagesAsync(
        ZipArchive templateArchive,
        OpenXmlZip outputArchive,
        Dictionary<string, (string DrawingPath, string DrawingRelsPath)> templateDrawings,
        HashSet<string> templateSheetRels,
        CancellationToken cancellationToken = default)
    {
        var imageFiles = _files.Where(file => file.IsImage).ToList();
        var wrappedDrawings = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var templateDrawingPaths = new HashSet<string>(templateDrawings.Values.Select(value => value.DrawingPath), StringComparer.OrdinalIgnoreCase);
        var writtenSheetRels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var sheetGroup in imageFiles.GroupBy(file => file.SheetIndex))
        {
            var sheetIndex = sheetGroup.Key;
            var files = sheetGroup.ToList();

            if (_sheetTemplateDrawings.TryGetValue(sheetIndex, out var templateDrawing))
            {
                wrappedDrawings.Add(templateDrawing.DrawingPath);
                await MergeIntoExistingDrawingAsync(templateArchive, outputArchive, templateDrawing, files, cancellationToken).ConfigureAwait(false);
            }
            else if (!templateDrawingPaths.Contains(ExcelFileNames.Drawing(sheetIndex)))
            {
                await EmitNewDrawingAsync(outputArchive, sheetIndex, files, cancellationToken).ConfigureAwait(false);
            }

            // Worksheet relationships: merge the drawing relationship into the template's rels, or
            // create a fresh rels part. Without this the <drawing r:id> would dangle and Excel would
            // repair the workbook by dropping the drawing.
            var sheetRelsPath = ExcelFileNames.SheetRels(sheetIndex);
            if (_sheetTemplateRels.TryGetValue(sheetIndex, out var templateRelsPath))
            {
                var relsDoc = await LoadXmlAsync(templateArchive, templateRelsPath, cancellationToken).ConfigureAwait(false);
                EnsureDrawingRelationship(relsDoc, sheetIndex);
                await SaveXmlToZipAsync(outputArchive.ZipFile, sheetRelsPath, relsDoc, cancellationToken).ConfigureAwait(false);
                writtenSheetRels.Add(templateRelsPath);
            }
            else
            {
                await WriteTextEntryAsync(outputArchive.ZipFile, sheetRelsPath, ExcelXml.DefaultSheetRelXml(ExcelXml.DrawingRelationship(sheetIndex)), cancellationToken).ConfigureAwait(false);
            }
        }

        // Template parts we deliberately did not copy must be written back when they were not reused.
        foreach (var relsPath in templateSheetRels)
        {
            if (!writtenSheetRels.Contains(relsPath))
                await CopyEntryAsync(templateArchive, outputArchive.ZipFile, relsPath, cancellationToken).ConfigureAwait(false);
        }

        foreach (var templateDrawing in templateDrawings.Values.Distinct())
        {
            if (wrappedDrawings.Contains(templateDrawing.DrawingPath))
                continue;

            await CopyEntryAsync(templateArchive, outputArchive.ZipFile, templateDrawing.DrawingPath, cancellationToken).ConfigureAwait(false);
            if (templateArchive.GetEntry(templateDrawing.DrawingRelsPath) is not null)
                await CopyEntryAsync(templateArchive, outputArchive.ZipFile, templateDrawing.DrawingRelsPath, cancellationToken).ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    private async Task EmitNewDrawingAsync(OpenXmlZip outputArchive, int sheetIndex, IReadOnlyList<FileDto> files, CancellationToken cancellationToken)
    {
        var drawingPath = ExcelFileNames.Drawing(sheetIndex);
        _createdDrawingParts.Add(drawingPath);

        var anchors = new StringBuilder();
        var drawingRels = new StringBuilder();

        var index = 0;
        foreach (var file in files)
        {
            await WriteBinaryEntryAsync(outputArchive.ZipFile, file.Path, file.Contents, cancellationToken).ConfigureAwait(false);
            anchors.Append(ExcelXml.DrawingXml(file, index));
            index++;
            drawingRels.AppendLine(ExcelXml.ImageRelationship(file));
        }

        await WriteTextEntryAsync(outputArchive.ZipFile, drawingPath, ExcelXml.DefaultDrawing(anchors.ToString()), cancellationToken).ConfigureAwait(false);
        await WriteTextEntryAsync(outputArchive.ZipFile, ExcelFileNames.DrawingRels(sheetIndex), ExcelXml.DefaultDrawingXmlRels(drawingRels.ToString()), cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private static async Task MergeIntoExistingDrawingAsync(
        ZipArchive templateArchive,
        OpenXmlZip outputArchive,
        (string DrawingPath, string DrawingRelsPath) templateDrawing,
        IReadOnlyList<FileDto> files,
        CancellationToken cancellationToken)
    {
        var drawingDoc = await LoadXmlAsync(templateArchive, templateDrawing.DrawingPath, cancellationToken).ConfigureAwait(false);
        var drawingRoot = drawingDoc.Root;
        if (drawingRoot is null)
            return;

        var maxPictureId = drawingRoot.Descendants()
            .Where(element => element.Name.LocalName == "cNvPr")
            .Select(element => int.TryParse(element.Attribute("id")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id) ? id : 0)
            .DefaultIfEmpty(0)
            .Max();

        var anchors = new StringBuilder();
        var drawingRels = new StringBuilder();
        var index = 0;
        foreach (var file in files)
        {
            await WriteBinaryEntryAsync(outputArchive.ZipFile, file.Path, file.Contents, cancellationToken).ConfigureAwait(false);
            anchors.Append(ExcelXml.DrawingXml(file, maxPictureId + index));
            index++;
            drawingRels.AppendLine(ExcelXml.ImageRelationship(file));
        }

        var ourAnchors = XDocument.Parse(ExcelXml.DefaultDrawing(anchors.ToString()));
        if (ourAnchors.Root is not null)
        {
            foreach (var anchor in ourAnchors.Root.Elements())
                drawingRoot.Add(anchor);
        }

        await SaveXmlToZipAsync(outputArchive.ZipFile, templateDrawing.DrawingPath, drawingDoc, cancellationToken).ConfigureAwait(false);

        var relsDoc = templateArchive.GetEntry(templateDrawing.DrawingRelsPath) is not null
            ? await LoadXmlAsync(templateArchive, templateDrawing.DrawingRelsPath, cancellationToken).ConfigureAwait(false)
            : XDocument.Parse(ExcelXml.DefaultDrawingXmlRels(string.Empty));

        var ourRels = XDocument.Parse(ExcelXml.DefaultDrawingXmlRels(drawingRels.ToString()));
        if (relsDoc.Root is not null && ourRels.Root is not null)
        {
            foreach (var relationship in ourRels.Root.Elements())
                relsDoc.Root.Add(relationship);
        }

        await SaveXmlToZipAsync(outputArchive.ZipFile, templateDrawing.DrawingRelsPath, relsDoc, cancellationToken).ConfigureAwait(false);
    }

    private static void EnsureDrawingRelationship(XDocument relsDoc, int sheetIndex)
    {
        var root = relsDoc.Root;
        if (root is null)
            return;

        var hasDrawingRelationship = root.Elements()
            .Any(element => element.Attribute("Type")?.Value == Schemas.SpreadsheetmlXmlDrawingRelationship);

        if (hasDrawingRelationship)
            return;

        var drawingRelationship = XDocument.Parse(ExcelXml.DefaultSheetRelXml(ExcelXml.DrawingRelationship(sheetIndex)));
        if (drawingRelationship.Root is not null)
        {
            foreach (var relationship in drawingRelationship.Root.Elements())
                root.Add(relationship);
        }
    }

    [CreateSyncVersion]
    private static async Task CopyEntryAsync(ZipArchive templateArchive, ZipArchive outputArchive, string path, CancellationToken cancellationToken)
    {
        var sourceEntry = templateArchive.GetEntry(path);
        if (sourceEntry is null)
            return;

        var targetEntry = outputArchive.CreateEntry(path);
        var sourceStream = await sourceEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableSource = sourceStream.ConfigureAwait(false);
        var targetStream = await targetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableTarget = targetStream.ConfigureAwait(false);
        await sourceStream.CopyToAsync(targetStream
#if NET
            , cancellationToken
#endif
        ).ConfigureAwait(false);
    }

    /// <summary>
    /// Ensures the persisted [Content_Types].xml declares a Default entry for every emitted image
    /// extension and an Override for every newly created drawing part.
    /// </summary>
    private void EnsureImageContentTypes(XDocument contentTypesDoc)
    {
        var root = contentTypesDoc.Root;
        if (root is null)
            return;

        var ns = root.Name.Namespace;
        foreach (var extension in _files.Where(file => file.IsImage).Select(file => file.Extension).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            var alreadyDeclared = root.Elements(ns + "Default")
                .Any(element => string.Equals(element.Attribute("Extension")?.Value, extension, StringComparison.OrdinalIgnoreCase));

            if (!alreadyDeclared)
            {
                root.Add(new XElement(ns + "Default",
                    new XAttribute("Extension", extension),
                    new XAttribute("ContentType", GetImageContentType(extension))));
            }
        }

        foreach (var drawingPath in _createdDrawingParts)
        {
            var partName = "/" + drawingPath;
            var alreadyDeclared = root.Elements(ns + "Override")
                .Any(element => string.Equals(element.Attribute("PartName")?.Value, partName, StringComparison.OrdinalIgnoreCase));

            if (!alreadyDeclared)
            {
                root.Add(new XElement(ns + "Override",
                    new XAttribute("PartName", partName),
                    new XAttribute("ContentType", ExcelContentTypes.Drawing)));
            }
        }
    }

    private static string GetImageContentType(string extension) => extension.ToLowerInvariant() switch
    {
        "png" => "image/png",
        "jpg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tiff" => "image/tiff",
        _ => "application/octet-stream"
    };

    [CreateSyncVersion]
    private static async Task WriteBinaryEntryAsync(ZipArchive zip, string path, byte[] contents, CancellationToken cancellationToken)
    {
        var entry = zip.CreateEntry(path);
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);
#if NET
        await stream.WriteAsync(contents.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
        await stream.WriteAsync(contents, 0, contents.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    [CreateSyncVersion]
    private static async Task WriteTextEntryAsync(ZipArchive zip, string path, string content, CancellationToken cancellationToken)
    {
        var entry = zip.CreateEntry(path);
        var stream = await entry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = stream.ConfigureAwait(false);
        var bytes = Encoding.UTF8.GetBytes(content);
#if NET
        await stream.WriteAsync(bytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
        await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }
}
