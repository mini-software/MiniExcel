namespace MiniExcelLib.OpenXml.Templates;

/// <summary>
/// Template image support: byte[] values resolved from template placeholders are emitted as embedded
/// images, reusing the same OpenXML primitives as the SaveAs pipeline (ImageHelper format detection,
/// FileDto and the ExcelXml drawing builders). See issue #972 / #604.
/// </summary>
internal partial class OpenXmlTemplate
{
    private const string ImageMarkerPrefix = "@@@imageid@@@,";
    private const long EmuPerPoint = 12700;

#if NET
    [GeneratedRegex($@"{ImageMarkerPrefix}(\d+)")]
    private static partial Regex ImageMarkerCellRegex();
    private static readonly Regex ImageMarkerCellRegexImpl = ImageMarkerCellRegex();
#else
    private static readonly Regex ImageMarkerCellRegexImpl = new($@"{ImageMarkerPrefix}(\d+)", RegexOptions.Compiled);
#endif
    
    private readonly List<FileDto> _files = [];
    private readonly Dictionary<string, PendingImage> _pendingImages = [];
    private readonly Dictionary<string, FileDto> _capturedImages = [];
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

    private void ResetImageState()
    {
        _files.Clear();
        _pendingImages.Clear();
        _capturedImages.Clear();
        _sheetTemplateDrawings.Clear();
        _sheetTemplateRels.Clear();
        _createdDrawingParts.Clear();
        _nextImageId = 0;
        _currentSheetIndex = 0;
    }

    /// <summary>
    /// Releases the per-sheet image bookkeeping once a worksheet has been rendered. Image bytes that
    /// were never captured (values resolved speculatively but not emitted) are dropped here, and
    /// captured entries are no longer needed because markers never cross worksheets.
    /// </summary>
    private void ReleaseSheetImageState()
    {
        _pendingImages.Clear();
        _capturedImages.Clear();
    }

    /// <summary>
    /// Clears the per-run image state when the template call ends, so a reused templater does not keep
    /// the last run's image bytes alive.
    /// </summary>
    private sealed class ImageStateScope(OpenXmlTemplate template) : IDisposable
    {
        public void Dispose() => template.ResetImageState();
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

        return $"{ImageMarkerPrefix}{id}";
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
            if (property is { CanRead: true } && property.GetIndexParameters().Length == 0)
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

    /// <summary>
    /// Replaces image markers embedded in a rendered row with empty cells, registering each image
    /// against its final cell reference. Coordinates are taken from the cell reference itself, which
    /// the surrounding code has already rewritten to the final (post collection expansion) row.
    /// </summary>
    private string CaptureAndClearImageMarkers(string rowXml, int sheetIndex)
    {
        if ((_pendingImages.Count == 0 && _capturedImages.Count == 0) || !rowXml.Contains(ImageMarkerPrefix))
            return rowXml;

        var rowElement = XElement.Parse(rowXml);

        var ht = rowElement.Attribute("ht")?.Value;
        var rowHeightPoints = double.TryParse(ht, NumberStyles.Float, CultureInfo.InvariantCulture, out var height)
            ? height 
            : 0;

        var colElements = rowElement.Elements(SpreadsheetNs  + "c")
            .Where(c => c.HasElements && c.Value.Contains(ImageMarkerPrefix));

        foreach (var col in colElements)
        {
            var cellRef = col.Attribute("r")?.Value;
            if (!CellReferenceConverter.TryParseCellReference(cellRef, out var column, out var row))
                continue; // invalid cell reference

            var textElement = col.Elements().First();
            if (textElement.HasElements)
                textElement = textElement.Elements().First();

            foreach (Match match in ImageMarkerCellRegexImpl.Matches(textElement.Value))
            {
                var imgId = match.Groups[1].Value;
                if (TryResolvePendingImage(imgId, out var pending))
                {
                    var file = new FileDto
                    {
                        SheetIndex = sheetIndex,
                        RowIndex = row,
                        CellIndex = column,
                        Contents = pending.Bytes,
                        Extension = pending.Extension,
                        IsImage = true,
        
                        // Two images can share the same anchor cell (two placeholders in one cell, or the
                        // same placeholder repeated), which would otherwise derive the same media part and
                        // relationship id. The per-file suffix keeps every derived identifier unique.
                        IdSuffix = (_files.Count + 1).ToString(CultureInfo.InvariantCulture)
                    };
        
                    ApplyRowHeightSize(file, pending, rowHeightPoints);
                    _files.Add(file);
                    _capturedImages[imgId] = file;
                }
            }

            var newCellValue = ImageMarkerCellRegexImpl.Replace(textElement.Value, "");
            textElement.SetValue(newCellValue.Trim());
        }

        return rowElement.ToString();
    }

    /// <summary>
    /// Resolves a pending marker to its image. The bytes are owned by <c>_pendingImages</c> until the
    /// first capture and are then released to the created <see cref="FileDto"/>; repeated captures
    /// (for example a grouped row rendered several times) reuse them through <c>_capturedImages</c>
    /// rather than keeping a second copy alive.
    /// </summary>
    private bool TryResolvePendingImage(string id, out PendingImage pending)
    {
        if (_pendingImages.TryGetValue(id, out var registered))
        {
            _pendingImages.Remove(id);
            pending = registered;
            return true;
        }

        if (_capturedImages.TryGetValue(id, out var captured))
        {
            pending = new PendingImage(captured.Contents, captured.Extension, ImageHelper.GetImageSize(captured.Contents));
            return true;
        }

        pending = null!;
        return false;
    }

    /// <summary>
    /// Sizes an image to the height of the row it is anchored to, preserving its aspect ratio. Rows
    /// without an explicit height keep the default anchor size.
    /// </summary>
    private static void ApplyRowHeightSize(FileDto file, PendingImage pending, double rowHeightPoints)
    {
        if (pending.Size is not { Width: > 0, Height: > 0 } size || rowHeightPoints <= 0)
            return;

        var heightEmu = (long)Math.Round(rowHeightPoints * EmuPerPoint);
        var widthEmu = (long)Math.Round(heightEmu * (size.Width / (double)size.Height));
        file.ImageWidthEmu = widthEmu;
        file.ImageHeightEmu = heightEmu;
    }

    [CreateSyncVersion]
    private static async Task WriteDrawingReferenceAsync(XmlWriter writer, string? prefix, int sheetIndex)
    {
        // Use the writer's namespace tracking instead of a raw string so that the r prefix referenced
        // by r:id is declared on the element whenever the worksheet does not already declare it.
        await writer.WriteStartElementAsync(prefix, "drawing", Schemas.SpreadsheetmlXmlMain).ConfigureAwait(false);
        await writer.WriteAttributeStringAsync("r", "id", Schemas.SpreadsheetmlXmlRelationships, $"rDrawing{sheetIndex}").ConfigureAwait(false);
        await writer.WriteEndElementAsync().ConfigureAwait(false);
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
        var writtenSheetRels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Drawing parts that already exist in the template keep their filename even when they belong
        // to a different sheet, so a generated drawing must never reuse one of them.
        var occupiedDrawingParts = GetTemplateDrawingPartNames(templateArchive);

        foreach (var sheetGroup in imageFiles.GroupBy(file => file.SheetIndex))
        {
            var sheetIndex = sheetGroup.Key;
            var files = sheetGroup.ToList();

            string drawingFileName;
            if (_sheetTemplateDrawings.TryGetValue(sheetIndex, out var templateDrawing))
            {
                wrappedDrawings.Add(templateDrawing.DrawingPath);
                await MergeIntoExistingDrawingAsync(templateArchive, outputArchive, templateDrawing, files, cancellationToken).ConfigureAwait(false);
                drawingFileName = Path.GetFileName(templateDrawing.DrawingPath);
            }
            else
            {
                // ExcelFileNames.Drawing(sheetIndex) may already be taken by an unrelated template
                // sheet, so allocate the first free deterministic drawing part name instead.
                var drawingPath = AllocateDrawingPart(sheetIndex, occupiedDrawingParts);
                occupiedDrawingParts.Add(drawingPath);
                drawingFileName = Path.GetFileName(drawingPath);
                await EmitNewDrawingAsync(outputArchive, drawingPath, files, cancellationToken).ConfigureAwait(false);
            }

            // Worksheet relationships: merge the drawing relationship into the template's rels, or
            // create a fresh rels part. Without this the <drawing r:id> would dangle and Excel would
            // repair the workbook by dropping the drawing. The relationship id keeps the per-sheet
            // convention while its target points at the drawing part allocated above.
            var sheetRelsPath = ExcelFileNames.SheetRels(sheetIndex);
            if (_sheetTemplateRels.TryGetValue(sheetIndex, out var templateRelsPath))
            {
                var relsDoc = await LoadXmlAsync(templateArchive, templateRelsPath, cancellationToken).ConfigureAwait(false);
                EnsureDrawingRelationship(relsDoc, sheetIndex, drawingFileName);
                await SaveXmlToZipAsync(outputArchive.ZipFile, sheetRelsPath, relsDoc, cancellationToken).ConfigureAwait(false);
                writtenSheetRels.Add(templateRelsPath);
            }
            else
            {
                await WriteTextEntryAsync(outputArchive.ZipFile, sheetRelsPath, ExcelXml.DefaultSheetRelXml(ExcelXml.DrawingRelationship(sheetIndex, drawingFileName)), cancellationToken).ConfigureAwait(false);
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
    private async Task EmitNewDrawingAsync(OpenXmlZip outputArchive, string drawingPath, IReadOnlyList<FileDto> files, CancellationToken cancellationToken)
    {
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
        await WriteTextEntryAsync(outputArchive.ZipFile, ExcelFileNames.DrawingRels(Path.GetFileName(drawingPath)), ExcelXml.DefaultDrawingXmlRels(drawingRels.ToString()), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Collects the drawing part names already present in the template so generated drawings never
    /// overwrite a part that belongs to another sheet.
    /// </summary>
    private static HashSet<string> GetTemplateDrawingPartNames(ZipArchive templateArchive)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in templateArchive.Entries)
        {
            var name = entry.FullName.TrimStart('/');
            if (name.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase) &&
                name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                names.Add(name);
            }
        }

        return names;
    }

    /// <summary>
    /// Returns the deterministic drawing part name to use for a generated sheet, preferring
    /// <c>drawing{sheetIndex}.xml</c> and only falling back to the first free number when that name is
    /// already taken by a template drawing part.
    /// </summary>
    private static string AllocateDrawingPart(int sheetIndex, ISet<string> occupied)
    {
        var preferred = ExcelFileNames.Drawing(sheetIndex);
        if (!occupied.Contains(preferred))
            return preferred;

        for (var candidateIndex = 1; ; candidateIndex++)
        {
            var candidate = ExcelFileNames.Drawing(candidateIndex);
            if (!occupied.Contains(candidate))
                return candidate;
        }
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

    private static void EnsureDrawingRelationship(XDocument relsDoc, int sheetIndex, string drawingFileName)
    {
        var root = relsDoc.Root;
        if (root is null)
            return;

        var hasDrawingRelationship = root.Elements()
            .Any(element => element.Attribute("Type")?.Value == Schemas.SpreadsheetmlXmlDrawingRelationship);

        if (hasDrawingRelationship)
            return;

        var drawingRelationship = XDocument.Parse(ExcelXml.DefaultSheetRelXml(ExcelXml.DrawingRelationship(sheetIndex, drawingFileName)));
        if (drawingRelationship.Root is not null)
        {
            foreach (var relationship in drawingRelationship.Root.Elements())
                root.Add(relationship);
        }
    }

    [CreateSyncVersion]
    private static async Task CopyEntryAsync(ZipArchive templateArchive, ZipArchive outputArchive, string path, CancellationToken cancellationToken)
    {
        if (templateArchive.GetEntry(path) is not { } sourceEntry)
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
                var imageContentType = extension switch
                {
                    "png" => "image/png",
                    "jpg" => "image/jpeg",
                    "gif" => "image/gif",
                    "bmp" => "image/bmp",
                    "tiff" => "image/tiff",
                    _ => "application/octet-stream"
                };

                root.Add(new XElement(ns + "Default",
                    new XAttribute("Extension", extension),
                    new XAttribute("ContentType", imageContentType)));
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
