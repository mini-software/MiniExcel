using MiniExcelLib.OpenXml.Models;

namespace MiniExcelLib.OpenXml.Picture;

/// <summary>
/// Writes Excel 365 "Place in Cell" images via workbook richData (not drawings).
/// </summary>
internal static partial class OpenXmlPlaceInCellImplement
{
    private const string SpreadsheetMlNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string PackageRelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string RichDataNs = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata";
    private const string RichData2Ns = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2";
    private const string RichValueRelNs = "http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel";
    private const string RelSheetMetadata = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
    private const string RelRdRichValue = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue";
    private const string RelRdRichValueStructure = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure";
    private const string RelRdRichValueTypes = "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes";
    private const string RelRichValueRel = "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel";
    private const string RelImage = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

    private const string MetadataPath = "xl/metadata.xml";
    private const string RdRichValuePath = "xl/richData/rdrichvalue.xml";
    private const string RdRichValueStructurePath = "xl/richData/rdrichvaluestructure.xml";
    private const string RdRichValueTypesPath = "xl/richData/rdRichValueTypes.xml";
    private const string RichValueRelPath = "xl/richData/richValueRel.xml";
    private const string RichValueRelRelsPath = "xl/richData/_rels/richValueRel.xml.rels";
    private const string WorkbookRelsPath = "xl/_rels/workbook.xml.rels";
    private const string ContentTypesPath = "[Content_Types].xml";

    private const string RvbExtUri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}";

    [CreateSyncVersion]
    public static async Task AddAsync(
        ZipArchive archive,
        IReadOnlyList<SheetRecord> sheetEntries,
        IReadOnlyList<MiniExcelPicture> images,
        CancellationToken cancellationToken = default)
    {
        if (images.Count == 0)
            return;

        if (sheetEntries.Count == 0)
            throw new InvalidOperationException("Workbook has no worksheets.");

        // Validate once up-front; collect content-type defaults without touching the zip repeatedly.
        var imageFormats = new HashSet<(string Ext, string ContentType)>();
        for (var i = 0; i < images.Count; i++)
        {
            var image = images[i];
            if (image.ImageBytes is null || image.ImageBytes.Length == 0)
                throw new InvalidDataException($"ImageBytes is required for PlaceInCell image at index {i}.");
            if (string.IsNullOrWhiteSpace(image.CellAddress))
                throw new InvalidDataException($"CellAddress is required for PlaceInCell image at index {i}.");
            if (!CellReferenceConverter.TryParseCellReference(image.CellAddress, out _, out _))
                throw new InvalidDataException($"Value {image.CellAddress} is not a valid cell reference.");

            imageFormats.Add(ResolveImageFormat(image.PictureType));
        }

        EnsureWorkbookRichDataRelationships(archive);
        EnsureContentTypes(archive, imageFormats);

        var structureDoc = LoadOrCreateXml(archive, RdRichValueStructurePath, CreateRichValueStructureXml);
        EnsureLocalImageStructure(structureDoc);
        var localImageStructureIndex = GetLocalImageStructureIndex(structureDoc);
        SaveXml(structureDoc, GetOrCreateEntry(archive, RdRichValueStructurePath));

        var typesEntry = archive.GetEntry(RdRichValueTypesPath);
        if (typesEntry is null)
            SaveXml(CreateRichValueTypesXml(), archive.CreateEntry(RdRichValueTypesPath));

        var rvDataDoc = LoadOrCreateXml(archive, RdRichValuePath, CreateRichValueDataXml);
        var richValueRelDoc = LoadOrCreateXml(archive, RichValueRelPath, CreateRichValueRelXml);
        var richValueRelsDoc = LoadOrCreateXml(archive, RichValueRelRelsPath, CreateRelationshipsXml);
        var metadataDoc = LoadOrCreateXml(archive, MetadataPath, CreateMetadataXml);

        var nextMediaIndex = GetNextMediaIndex(archive);
        var existingRelCount = CountChildElements(richValueRelDoc.DocumentElement, "rel");
        var existingRvCount = CountChildElements(rvDataDoc.DocumentElement, "rv");
        var nextRelIdNumber = GetMaxRelationshipIdNumber(richValueRelsDoc.DocumentElement) + 1;
        EnsureXlRichValueMetadataType(metadataDoc);
        var xlRichValueTypeIndex = GetMetadataTypeIndex(metadataDoc, "XLRICHVALUE");

        var sheetDocs = new Dictionary<string, (XmlDocument Doc, ZipArchiveEntry Entry, int MaxCol, int MaxRow)>(
            StringComparer.OrdinalIgnoreCase);

        for (var i = 0; i < images.Count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var image = images[i];
            var cellAddress = image.CellAddress!.ToUpperInvariant();
            CellReferenceConverter.TryParseCellReference(cellAddress, out var column, out var row);

            var (ext, _) = ResolveImageFormat(image.PictureType);
            var mediaFileName = $"image{nextMediaIndex++}.{ext}";
            var mediaPath = $"xl/media/{mediaFileName}";
            await WriteMediaAsync(archive, mediaPath, image.ImageBytes, cancellationToken).ConfigureAwait(false);

            var relIndex = existingRelCount + i; // 0-based LocalImageIdentifier
            var relId = $"rId{nextRelIdNumber++}";
            var rvIndex = existingRvCount + i; // 0-based
            var vm = AppendValueMetadata(metadataDoc, rvIndex, xlRichValueTypeIndex); // 1-based

            AppendRichValueRel(richValueRelDoc, relId);
            AppendRelationship(richValueRelsDoc, relId, RelImage, $"../media/{mediaFileName}");
            AppendRichValue(rvDataDoc, relIndex, localImageStructureIndex);

            var sheetEnt = sheetEntries.FirstOrDefault(x =>
                              string.Equals(x.Name, image.SheetName ?? sheetEntries[0].Name, StringComparison.OrdinalIgnoreCase))
                          ?? sheetEntries[0];
            var sheetPath = ResolveSheetPath(sheetEnt);

            if (!sheetDocs.TryGetValue(sheetPath, out var sheetInfo))
            {
                var sheetEntry = archive.GetEntry(sheetPath)
                                 ?? throw new InvalidOperationException($"Worksheet part '{sheetPath}' was not found.");
                sheetInfo = (LoadXml(sheetEntry), sheetEntry, column, row);
                sheetDocs[sheetPath] = sheetInfo;
            }
            else
            {
                sheetInfo = (
                    sheetInfo.Doc,
                    sheetInfo.Entry,
                    Math.Max(sheetInfo.MaxCol, column),
                    Math.Max(sheetInfo.MaxRow, row));
                sheetDocs[sheetPath] = sheetInfo;
            }

            // Overwrites any existing value/formula in the target cell (Place in Cell is the cell value).
            UpsertImageCell(sheetInfo.Doc, cellAddress, vm);
        }

        SaveXml(rvDataDoc, GetOrCreateEntry(archive, RdRichValuePath));
        SaveXml(richValueRelDoc, GetOrCreateEntry(archive, RichValueRelPath));
        SaveXml(richValueRelsDoc, GetOrCreateEntry(archive, RichValueRelRelsPath));
        SaveXml(metadataDoc, GetOrCreateEntry(archive, MetadataPath));

        foreach (var pair in sheetDocs)
        {
            ExpandDimension(pair.Value.Doc, pair.Value.MaxCol, pair.Value.MaxRow);
            SaveXml(pair.Value.Doc, pair.Value.Entry);
        }
    }

    private static string ResolveSheetPath(SheetRecord sheetEnt)
    {
        if (!string.IsNullOrEmpty(sheetEnt.Path))
        {
            var path = sheetEnt.Path.Replace('\\', '/');
            if (path.StartsWith("/", StringComparison.Ordinal))
                path = path.Substring(1);
            if (path.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                return path;
            return $"xl/{path.TrimStart('/')}";
        }

        return $"xl/worksheets/sheet{sheetEnt.Id}.xml";
    }

    [CreateSyncVersion]
    private static async Task WriteMediaAsync(ZipArchive archive, string mediaPath, byte[] imageBytes, CancellationToken cancellationToken)
    {
        var imageEntry = archive.GetEntry(mediaPath) ?? archive.CreateEntry(mediaPath);
#if NET8_0_OR_GREATER
        var entryStream = await imageEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableStream = entryStream.ConfigureAwait(false);
        entryStream.SetLength(0);
        await entryStream.WriteAsync(imageBytes.AsMemory(), cancellationToken).ConfigureAwait(false);
#else
        using var entryStream = imageEntry.Open();
        entryStream.SetLength(0);
        await entryStream.WriteAsync(imageBytes, 0, imageBytes.Length, cancellationToken).ConfigureAwait(false);
#endif
    }

    private static (string Ext, string ContentType) ResolveImageFormat(string? pictureType)
    {
        var type = (pictureType ?? "png").Trim().ToLowerInvariant();
        if (type.StartsWith("image/", StringComparison.Ordinal))
            type = type.Substring("image/".Length);

        return type switch
        {
            "jpg" or "jpeg" => ("jpeg", "image/jpeg"),
            "gif" => ("gif", "image/gif"),
            "bmp" => ("bmp", "image/bmp"),
            "webp" => ("webp", "image/webp"),
            "tif" or "tiff" => ("tiff", "image/tiff"),
            _ => ("png", "image/png")
        };
    }

    private static int GetNextMediaIndex(ZipArchive archive)
    {
        var max = 0;
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.StartsWith("xl/media/image", StringComparison.OrdinalIgnoreCase))
                continue;

            var name = Path.GetFileNameWithoutExtension(entry.FullName);
            if (name.StartsWith("image", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(name.Substring("image".Length), out var index))
            {
                max = Math.Max(max, index);
            }
        }

        return max + 1;
    }

    private static void EnsureWorkbookRichDataRelationships(ZipArchive archive)
    {
        var entry = GetOrCreateEntry(archive, WorkbookRelsPath);
        var doc = LoadOrCreateXml(archive, WorkbookRelsPath, CreateRelationshipsXml);
        var root = doc.DocumentElement
                   ?? throw new InvalidOperationException("workbook.xml.rels is missing a root element.");

        EnsureRelationship(root, RelSheetMetadata, "metadata.xml");
        EnsureRelationship(root, RelRdRichValue, "richData/rdrichvalue.xml");
        EnsureRelationship(root, RelRdRichValueStructure, "richData/rdrichvaluestructure.xml");
        EnsureRelationship(root, RelRdRichValueTypes, "richData/rdRichValueTypes.xml");
        EnsureRelationship(root, RelRichValueRel, "richData/richValueRel.xml");

        SaveXml(doc, entry);
    }

    private static void EnsureRelationship(XmlElement root, string type, string target)
    {
        foreach (XmlElement rel in root.ChildNodes.OfType<XmlElement>())
        {
            if (string.Equals(rel.GetAttribute("Type"), type, StringComparison.Ordinal)
                && string.Equals(rel.GetAttribute("Target"), target, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        var nextId = GetNextRelationshipId(root);
        var node = root.OwnerDocument!.CreateElement("Relationship", PackageRelsNs);
        node.SetAttribute("Id", nextId);
        node.SetAttribute("Type", type);
        node.SetAttribute("Target", target);
        root.AppendChild(node);
    }

    private static string GetNextRelationshipId(XmlElement root)
        => $"rId{GetMaxRelationshipIdNumber(root) + 1}";

    /// <summary>Returns the highest numeric suffix among rIdN relationships (0 if none).</summary>
    private static int GetMaxRelationshipIdNumber(XmlElement? root)
    {
        var max = 0;
        if (root is null)
            return max;

        foreach (XmlElement rel in root.ChildNodes.OfType<XmlElement>())
        {
            var id = rel.GetAttribute("Id");
            if (id.StartsWith("rId", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(id.Substring(3), out var n))
            {
                max = Math.Max(max, n);
            }
        }

        return max;
    }

    private static void EnsureContentTypes(ZipArchive archive, HashSet<(string Ext, string ContentType)> imageFormats)
    {
        var entry = archive.GetEntry(ContentTypesPath)
                    ?? throw new InvalidOperationException("[Content_Types].xml was not found.");
        var doc = LoadXml(entry);
        var root = doc.DocumentElement
                   ?? throw new InvalidOperationException("[Content_Types].xml is missing a root element.");

        EnsureOverride(root, "/xl/metadata.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
        EnsureOverride(root, "/xl/richData/rdrichvalue.xml", "application/vnd.ms-excel.rdrichvalue+xml");
        EnsureOverride(root, "/xl/richData/rdrichvaluestructure.xml", "application/vnd.ms-excel.rdrichvaluestructure+xml");
        EnsureOverride(root, "/xl/richData/rdRichValueTypes.xml", "application/vnd.ms-excel.rdrichvaluetypes+xml");
        EnsureOverride(root, "/xl/richData/richValueRel.xml", "application/vnd.ms-excel.richvaluerel+xml");

        // Single pass over Defaults for all image extensions used in this batch.
        var existingDefaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (XmlElement child in root.ChildNodes.OfType<XmlElement>())
        {
            if (child.LocalName == "Default")
                existingDefaults.Add(child.GetAttribute("Extension"));
        }

        foreach (var (ext, contentType) in imageFormats)
        {
            if (existingDefaults.Contains(ext))
                continue;

            var node = doc.CreateElement("Default", ContentTypesNs);
            node.SetAttribute("Extension", ext);
            node.SetAttribute("ContentType", contentType);
            root.AppendChild(node);
            existingDefaults.Add(ext);
        }

        // One save for Overrides + any new image Defaults.
        SaveXml(doc, entry);
    }

    private static void EnsureOverride(XmlElement root, string partName, string contentType)
    {
        foreach (XmlElement child in root.ChildNodes.OfType<XmlElement>())
        {
            if (child.LocalName == "Override"
                && string.Equals(child.GetAttribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }
        }

        var node = root.OwnerDocument!.CreateElement("Override", ContentTypesNs);
        node.SetAttribute("PartName", partName);
        node.SetAttribute("ContentType", contentType);
        root.AppendChild(node);
    }

    private static void EnsureLocalImageStructure(XmlDocument structureDoc)
    {
        var root = structureDoc.DocumentElement
                   ?? throw new InvalidOperationException("rdrichvaluestructure.xml is missing a root element.");

        var nsMgr = CreateNs(structureDoc, "xlrd", RichDataNs);
        if (root.SelectSingleNode("xlrd:s[@t='_localImage']", nsMgr) is not null)
            return;

        var s = structureDoc.CreateElement("s", RichDataNs);
        s.SetAttribute("t", "_localImage");

        var k1 = structureDoc.CreateElement("k", RichDataNs);
        k1.SetAttribute("n", "_rvRel:LocalImageIdentifier");
        k1.SetAttribute("t", "i");
        s.AppendChild(k1);

        var k2 = structureDoc.CreateElement("k", RichDataNs);
        k2.SetAttribute("n", "CalcOrigin");
        k2.SetAttribute("t", "i");
        s.AppendChild(k2);

        root.AppendChild(s);
        root.SetAttribute("count", CountChildElements(root, "s").ToString());
    }

    private static void EnsureXlRichValueMetadataType(XmlDocument metadataDoc)
    {
        var root = metadataDoc.DocumentElement
                   ?? throw new InvalidOperationException("metadata.xml is missing a root element.");

        EnsureNamespace(root, "xlrd", RichDataNs);

        var nsMgr = CreateNs(metadataDoc, "x", SpreadsheetMlNs);
        var types = root.SelectSingleNode("x:metadataTypes", nsMgr) as XmlElement;
        if (types is null)
        {
            types = metadataDoc.CreateElement("metadataTypes", SpreadsheetMlNs);
            root.InsertBefore(types, root.FirstChild);
        }

        if (types.SelectSingleNode("x:metadataType[@name='XLRICHVALUE']", nsMgr) is null)
        {
            var type = metadataDoc.CreateElement("metadataType", SpreadsheetMlNs);
            type.SetAttribute("name", "XLRICHVALUE");
            type.SetAttribute("minSupportedVersion", "120000");
            type.SetAttribute("copy", "1");
            type.SetAttribute("pasteAll", "1");
            type.SetAttribute("pasteValues", "1");
            type.SetAttribute("merge", "1");
            type.SetAttribute("splitFirst", "1");
            type.SetAttribute("rowColShift", "1");
            type.SetAttribute("clearFormats", "1");
            type.SetAttribute("clearComments", "1");
            type.SetAttribute("assign", "1");
            type.SetAttribute("coerce", "1");
            types.AppendChild(type);
        }

        types.SetAttribute("count", CountChildElements(types, "metadataType").ToString());

        var future = root.SelectSingleNode("x:futureMetadata[@name='XLRICHVALUE']", nsMgr) as XmlElement;
        if (future is null)
        {
            future = metadataDoc.CreateElement("futureMetadata", SpreadsheetMlNs);
            future.SetAttribute("name", "XLRICHVALUE");
            future.SetAttribute("count", "0");

            var valueMetadata = root.SelectSingleNode("x:valueMetadata", nsMgr);
            if (valueMetadata is not null)
                root.InsertBefore(future, valueMetadata);
            else
                root.AppendChild(future);
        }

        if (root.SelectSingleNode("x:valueMetadata", nsMgr) is null)
        {
            var valueMetadata = metadataDoc.CreateElement("valueMetadata", SpreadsheetMlNs);
            valueMetadata.SetAttribute("count", "0");
            root.AppendChild(valueMetadata);
        }
    }

    private static int GetMetadataTypeIndex(XmlDocument metadataDoc, string name)
    {
        var nsMgr = CreateNs(metadataDoc, "x", SpreadsheetMlNs);
        var types = metadataDoc.SelectSingleNode("/x:metadata/x:metadataTypes", nsMgr) as XmlElement
                    ?? throw new InvalidOperationException("metadataTypes was not found.");

        var index = 0;
        foreach (XmlElement type in types.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == "metadataType"))
        {
            index++;
            if (string.Equals(type.GetAttribute("name"), name, StringComparison.Ordinal))
                return index;
        }

        throw new InvalidOperationException($"metadataType '{name}' was not found.");
    }

    /// <summary>Appends futureMetadata + valueMetadata and returns 1-based vm index.</summary>
    private static int AppendValueMetadata(XmlDocument metadataDoc, int rvIndex, int xlRichValueTypeIndex)
    {
        var nsMgr = CreateNs(metadataDoc, "x", SpreadsheetMlNs);
        nsMgr.AddNamespace("xlrd", RichDataNs);

        var future = metadataDoc.SelectSingleNode("/x:metadata/x:futureMetadata[@name='XLRICHVALUE']", nsMgr) as XmlElement
                     ?? throw new InvalidOperationException("futureMetadata XLRICHVALUE was not found.");
        var valueMetadata = metadataDoc.SelectSingleNode("/x:metadata/x:valueMetadata", nsMgr) as XmlElement
                            ?? throw new InvalidOperationException("valueMetadata was not found.");

        var xlrdPrefix = metadataDoc.DocumentElement!.GetPrefixOfNamespace(RichDataNs);
        if (string.IsNullOrEmpty(xlrdPrefix))
        {
            metadataDoc.DocumentElement.SetAttribute("xmlns:xlrd", RichDataNs);
            xlrdPrefix = "xlrd";
        }

        var bk = metadataDoc.CreateElement("bk", SpreadsheetMlNs);
        var extLst = metadataDoc.CreateElement("extLst", SpreadsheetMlNs);
        var ext = metadataDoc.CreateElement("ext", SpreadsheetMlNs);
        ext.SetAttribute("uri", RvbExtUri);
        var rvb = metadataDoc.CreateElement(xlrdPrefix, "rvb", RichDataNs);
        rvb.SetAttribute("i", rvIndex.ToString());
        ext.AppendChild(rvb);
        extLst.AppendChild(ext);
        bk.AppendChild(extLst);
        future.AppendChild(bk);
        future.SetAttribute("count", CountChildElements(future, "bk").ToString());

        var vmBk = metadataDoc.CreateElement("bk", SpreadsheetMlNs);
        var rc = metadataDoc.CreateElement("rc", SpreadsheetMlNs);
        rc.SetAttribute("t", xlRichValueTypeIndex.ToString());
        rc.SetAttribute("v", rvIndex.ToString());
        vmBk.AppendChild(rc);
        valueMetadata.AppendChild(vmBk);

        var vmCount = CountChildElements(valueMetadata, "bk");
        valueMetadata.SetAttribute("count", vmCount.ToString());
        return vmCount; // 1-based
    }

    private static void AppendRichValue(XmlDocument rvDataDoc, int relIndex, int localImageStructureIndex)
    {
        var root = rvDataDoc.DocumentElement!;
        var rv = rvDataDoc.CreateElement("rv", RichDataNs);
        rv.SetAttribute("s", localImageStructureIndex.ToString());

        var v1 = rvDataDoc.CreateElement("v", RichDataNs);
        v1.InnerText = relIndex.ToString();
        rv.AppendChild(v1);

        var v2 = rvDataDoc.CreateElement("v", RichDataNs);
        v2.InnerText = "5"; // CalcOrigin = Standalone
        rv.AppendChild(v2);

        root.AppendChild(rv);
        root.SetAttribute("count", CountChildElements(root, "rv").ToString());
    }

    private static int GetLocalImageStructureIndex(XmlDocument structureDoc)
    {
        var root = structureDoc.DocumentElement
                   ?? throw new InvalidOperationException("rdrichvaluestructure.xml is missing a root element.");

        var index = 0;
        foreach (var s in root.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == "s"))
        {
            if (string.Equals(s.GetAttribute("t"), "_localImage", StringComparison.Ordinal))
                return index;
            index++;
        }

        return 0;
    }

    private static void AppendRichValueRel(XmlDocument richValueRelDoc, string relId)
    {
        var root = richValueRelDoc.DocumentElement!;
        EnsureNamespace(root, "r", OfficeRelsNs);

        var rel = richValueRelDoc.CreateElement("rel", RichValueRelNs);
        rel.SetAttribute("id", OfficeRelsNs, relId);
        root.AppendChild(rel);
    }

    private static void AppendRelationship(XmlDocument relsDoc, string relId, string type, string target)
    {
        var root = relsDoc.DocumentElement!;
        var node = relsDoc.CreateElement("Relationship", PackageRelsNs);
        node.SetAttribute("Id", relId);
        node.SetAttribute("Type", type);
        node.SetAttribute("Target", target);
        root.AppendChild(node);
    }

    /// <summary>
    /// Writes an in-cell image placeholder into the sheet.
    /// Replaces any existing cell value/formula; keeps style index <c>s</c> when present.
    /// </summary>
    private static void UpsertImageCell(XmlDocument sheetDoc, string cellAddress, int vm)
    {
        var ns = sheetDoc.DocumentElement?.NamespaceURI ?? SpreadsheetMlNs;
        var nsMgr = CreateNs(sheetDoc, "x", ns);

        if (!CellReferenceConverter.TryParseCellReference(cellAddress, out _, out var rowNumber))
            throw new InvalidDataException($"Value {cellAddress} is not a valid cell reference.");

        var worksheet = sheetDoc.DocumentElement
                        ?? throw new InvalidOperationException("Worksheet root was not found.");

        var sheetData = worksheet.SelectSingleNode("x:sheetData", nsMgr) as XmlElement;
        if (sheetData is null)
        {
            sheetData = sheetDoc.CreateElement("sheetData", ns);
            InsertSheetData(worksheet, sheetData, nsMgr);
        }

        var row = FindOrCreateRow(sheetDoc, sheetData, rowNumber, ns);
        var cell = FindCell(row, cellAddress);

        if (cell is null)
        {
            cell = sheetDoc.CreateElement("c", ns);
            cell.SetAttribute("r", cellAddress);
            InsertCellInRow(row, cell, cellAddress);
        }

        var style = cell.GetAttribute("s");
        cell.RemoveAllAttributes();
        cell.SetAttribute("r", cellAddress);
        if (!string.IsNullOrEmpty(style))
            cell.SetAttribute("s", style);
        cell.SetAttribute("t", "e");
        cell.SetAttribute("vm", vm.ToString());

        while (cell.HasChildNodes)
            cell.RemoveChild(cell.FirstChild!);

        var v = sheetDoc.CreateElement("v", ns);
        v.InnerText = "#VALUE!";
        cell.AppendChild(v);
    }

    /// <summary>
    /// Expands worksheet dimension to include the given 1-based cell when needed (O(1)).
    /// </summary>
    private static void ExpandDimension(XmlDocument sheetDoc, int maxCol, int maxRow)
    {
        var ns = sheetDoc.DocumentElement?.NamespaceURI ?? SpreadsheetMlNs;
        var nsMgr = CreateNs(sheetDoc, "x", ns);
        var worksheet = sheetDoc.DocumentElement;
        if (worksheet is null)
            return;

        var dimension = worksheet.SelectSingleNode("x:dimension", nsMgr) as XmlElement;
        var endRef = CellReferenceConverter.GetCellFromCoordinates(maxCol, maxRow);

        if (dimension is null)
        {
            dimension = sheetDoc.CreateElement("dimension", ns);
            dimension.SetAttribute("ref", $"A1:{endRef}");
            var sheetData = worksheet.SelectSingleNode("x:sheetData", nsMgr);
            if (sheetData is not null)
                worksheet.InsertBefore(dimension, sheetData);
            else
                worksheet.AppendChild(dimension);
            return;
        }

        var current = dimension.GetAttribute("ref");
        if (string.IsNullOrEmpty(current))
        {
            dimension.SetAttribute("ref", $"A1:{endRef}");
            return;
        }

        var parts = current.Split(':');
        var start = parts[0];
        var end = parts.Length > 1 ? parts[1] : parts[0];

        if (!CellReferenceConverter.TryParseCellReference(start, out var startCol, out var startRow)
            || !CellReferenceConverter.TryParseCellReference(end, out var endCol, out var endRow))
        {
            dimension.SetAttribute("ref", $"A1:{endRef}");
            return;
        }

        var newEndCol = Math.Max(endCol, maxCol);
        var newEndRow = Math.Max(endRow, maxRow);
        if (newEndCol == endCol && newEndRow == endRow)
            return;

        var newStart = CellReferenceConverter.GetCellFromCoordinates(Math.Min(startCol, maxCol), Math.Min(startRow, maxRow));
        // Keep original start if image is within/after the existing range start
        if (startCol <= maxCol && startRow <= maxRow)
            newStart = start;

        var newEnd = CellReferenceConverter.GetCellFromCoordinates(newEndCol, newEndRow);
        dimension.SetAttribute("ref", newStart == newEnd ? newStart : $"{newStart}:{newEnd}");
    }

    private static void InsertSheetData(XmlElement worksheet, XmlElement sheetData, XmlNamespaceManager nsMgr)
    {
        // sheetData typically follows sheetFormatPr / cols
        var after = worksheet.SelectSingleNode("x:cols", nsMgr)
                    ?? worksheet.SelectSingleNode("x:sheetFormatPr", nsMgr)
                    ?? worksheet.SelectSingleNode("x:sheetViews", nsMgr)
                    ?? worksheet.SelectSingleNode("x:dimension", nsMgr)
                    ?? worksheet.SelectSingleNode("x:sheetPr", nsMgr);

        if (after?.NextSibling is not null)
            worksheet.InsertAfter(sheetData, after);
        else if (after is not null)
            worksheet.AppendChild(sheetData);
        else if (worksheet.FirstChild is not null)
            worksheet.InsertBefore(sheetData, worksheet.FirstChild);
        else
            worksheet.AppendChild(sheetData);
    }

    private static XmlElement FindOrCreateRow(XmlDocument sheetDoc, XmlElement sheetData, int rowNumber, string ns)
    {
        XmlElement? insertBefore = null;
        foreach (var row in sheetData.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == "row"))
        {
            if (!int.TryParse(row.GetAttribute("r"), out var r))
                continue;

            if (r == rowNumber)
                return row;
            if (r > rowNumber)
            {
                insertBefore = row;
                break;
            }
        }

        var newRow = sheetDoc.CreateElement("row", ns);
        newRow.SetAttribute("r", rowNumber.ToString());
        if (insertBefore is not null)
            sheetData.InsertBefore(newRow, insertBefore);
        else
            sheetData.AppendChild(newRow);
        return newRow;
    }

    private static XmlElement? FindCell(XmlElement row, string cellAddress)
    {
        foreach (var cell in row.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == "c"))
        {
            if (string.Equals(cell.GetAttribute("r"), cellAddress, StringComparison.OrdinalIgnoreCase))
                return cell;
        }

        return null;
    }

    private static void InsertCellInRow(XmlElement row, XmlElement cell, string cellAddress)
    {
        CellReferenceConverter.TryParseCellReference(cellAddress, out var column, out _);
        XmlElement? insertBefore = null;
        foreach (var existing in row.ChildNodes.OfType<XmlElement>().Where(e => e.LocalName == "c"))
        {
            if (!CellReferenceConverter.TryParseCellReference(existing.GetAttribute("r"), out var col, out _))
                continue;
            if (col > column)
            {
                insertBefore = existing;
                break;
            }
        }

        if (insertBefore is not null)
            row.InsertBefore(cell, insertBefore);
        else
            row.AppendChild(cell);
    }

    private static int CountChildElements(XmlElement? parent, string localName)
    {
        if (parent is null)
            return 0;
        return parent.ChildNodes.OfType<XmlElement>().Count(e => e.LocalName == localName);
    }

    private static void EnsureNamespace(XmlElement element, string prefix, string ns)
    {
        if (!string.IsNullOrEmpty(element.GetPrefixOfNamespace(ns)))
            return;
        element.SetAttribute($"xmlns:{prefix}", ns);
    }

    private static XmlNamespaceManager CreateNs(XmlDocument doc, string prefix, string ns)
    {
        var mgr = new XmlNamespaceManager(doc.NameTable);
        mgr.AddNamespace(prefix, ns);
        return mgr;
    }

    private static ZipArchiveEntry GetOrCreateEntry(ZipArchive archive, string path)
        => archive.GetEntry(path) ?? archive.CreateEntry(path);

    private static XmlDocument LoadOrCreateXml(ZipArchive archive, string path, Func<XmlDocument> factory)
    {
        var entry = archive.GetEntry(path);
        if (entry is null)
            return factory();

        // Length is unavailable after an entry was opened for write in Update mode.
        try
        {
            if (entry.Length == 0)
                return factory();
        }
        catch (InvalidOperationException)
        {
            // Fall through and try to load from the entry stream.
        }

        return LoadXml(entry);
    }

    private static XmlDocument LoadXml(ZipArchiveEntry entry)
    {
        var doc = new XmlDocument { PreserveWhitespace = false };
        using var stream = entry.Open();
        if (stream.Length == 0)
            return new XmlDocument();

        doc.Load(stream);
        return doc;
    }

    private static void SaveXml(XmlDocument doc, ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        stream.SetLength(0);
        using var writer = XmlWriter.Create(stream, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            Indent = false,
            OmitXmlDeclaration = false
        });
        doc.Save(writer);
    }

    private static XmlDocument CreateRelationshipsXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
        return doc;
    }

    private static XmlDocument CreateRichValueStructureXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="0"/>
            """);
        return doc;
    }

    private static XmlDocument CreateRichValueDataXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="0"/>
            """);
        return doc;
    }

    private static XmlDocument CreateRichValueRelXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
                           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
            """);
        return doc;
    }

    private static XmlDocument CreateMetadataXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                      xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
              <metadataTypes count="0"/>
              <valueMetadata count="0"/>
            </metadata>
            """);
        return doc;
    }

    private static XmlDocument CreateRichValueTypesXml()
    {
        var doc = new XmlDocument();
        doc.LoadXml(
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
                         xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                         mc:Ignorable="x"
                         xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <global>
                <keyFlags>
                  <key name="_Self">
                    <flag name="ExcludeFromFile" value="1"/>
                    <flag name="ExcludeFromCalcComparison" value="1"/>
                  </key>
                  <key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                  <key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key>
                </keyFlags>
              </global>
            </rvTypesInfo>
            """);
        return doc;
    }
}
