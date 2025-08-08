namespace MiniExcelLib.Core.OpenXml.Styles.Builder;

internal abstract partial class SheetStyleBuilderBase(SheetStyleBuildContext context) : ISheetStyleBuilder
{
    private readonly SheetStyleBuildContext _context = context;
    
    //todo: these may actually be null if used when the context is not initialized
    private XmlReader OldReader => _context.OldXmlReader!;
    private XmlWriter NewWriter => _context.NewXmlWriter!;
    
    internal static readonly Dictionary<string, int> AllElements = new()
    {
        ["numFmts"] = 0,
        ["fonts"] = 1,
        ["fills"] = 2,
        ["borders"] = 3,
        ["cellStyleXfs"] = 4,
        ["cellXfs"] = 5,
        ["cellStyles"] = 6,
        ["dxfs"] = 7,
        ["tableStyles"] = 8,
        ["extLst"] = 9
    };


    // Todo: add CancellationToken to all methods called inside of BuildAsync
    [CreateSyncVersion]
    public virtual async Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default)
    {
        await _context.InitializeAsync(GetGenerateElementInfos(), cancellationToken).ConfigureAwait(false);

        while (await OldReader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            switch (OldReader.NodeType)
            {
                case XmlNodeType.Element:
                    await GenerateElementBeforStartElementAsync().ConfigureAwait(false);
                    await NewWriter.WriteStartElementAsync(OldReader.Prefix, OldReader.LocalName, OldReader.NamespaceURI).ConfigureAwait(false);
                    await WriteAttributesAsync(OldReader.LocalName, cancellationToken).ConfigureAwait(false);
                    if (OldReader.IsEmptyElement)
                    {
                        await GenerateElementBeforEndElementAsync().ConfigureAwait(false);
                        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
                    }
                    break;
                case XmlNodeType.Text:
                    await NewWriter.WriteStringAsync(OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    await NewWriter.WriteWhitespaceAsync(OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.CDATA:
                    await NewWriter.WriteCDataAsync(OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.EntityReference:
                    await NewWriter.WriteEntityRefAsync(OldReader.Name).ConfigureAwait(false);
                    break;
                case XmlNodeType.XmlDeclaration:
                case XmlNodeType.ProcessingInstruction:
                    await NewWriter.WriteProcessingInstructionAsync(OldReader.Name, OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.DocumentType:
                    await NewWriter.WriteDocTypeAsync(OldReader.Name, OldReader.GetAttribute("PUBLIC"), OldReader.GetAttribute("SYSTEM"), OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.Comment:
                    await NewWriter.WriteCommentAsync(OldReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.EndElement:
                    await GenerateElementBeforEndElementAsync().ConfigureAwait(false);
                    await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
                    break;
            }
        }

        await _context.FinalizeAndUpdateZipDictionaryAsync(cancellationToken).ConfigureAwait(false);

        return new SheetStyleBuildResult(GetCellXfIdMap());
    }

    protected abstract SheetStyleElementInfos GetGenerateElementInfos();

    [CreateSyncVersion]
    protected virtual async Task WriteAttributesAsync(string element, CancellationToken cancellationToken = default)
    {
        if (OldReader.NodeType is XmlNodeType.Element or XmlNodeType.XmlDeclaration)
        {
            if (OldReader.MoveToFirstAttribute())
            {
                await WriteAttributesAsync(element, cancellationToken).ConfigureAwait(false);
                OldReader.MoveToElement();
            }
        }
        else if (OldReader.NodeType == XmlNodeType.Attribute)
        {
            do
            {
                NewWriter.WriteStartAttribute(OldReader.Prefix, OldReader.LocalName, OldReader.NamespaceURI);
                var currentAttribute = OldReader.LocalName;
                while (OldReader.ReadAttributeValue())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                        
                    if (OldReader.NodeType == XmlNodeType.EntityReference)
                    {
                        await NewWriter.WriteEntityRefAsync(OldReader.Name).ConfigureAwait(false);
                    }
                    else if (currentAttribute == "count")
                    {
                        var value = element switch
                        {
                            "numFmts" => (_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString(),
                            "fonts" => (_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString(),
                            "fills" => (_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString(),
                            "borders" => (_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString(),
                            "cellStyleXfs" => (_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString(),
                            "cellXfs" => (_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString(),
                            _ => OldReader.Value
                        };
                        await NewWriter.WriteStringAsync(value).ConfigureAwait(false);
                    }
                    else
                    {
                        await NewWriter.WriteStringAsync(OldReader.Value).ConfigureAwait(false);
                    }
                }
                NewWriter.WriteEndAttribute();
            }
            while (OldReader.MoveToNextAttribute());
        }
    }

    [CreateSyncVersion]
    protected virtual async Task GenerateElementBeforStartElementAsync()
    {
        if (!AllElements.TryGetValue(OldReader.LocalName, out var elementIndex))
            return;
        
        if (!_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts && AllElements["numFmts"] < elementIndex)
        {
            await GenerateNumFmtsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsNumFmts = true;
        }
        else if (!_context.OldElementInfos.ExistsFonts && !_context.GenerateElementInfos.ExistsFonts && AllElements["fonts"] < elementIndex)
        {
            await GenerateFontsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsFonts = true;
        }
        else if (!_context.OldElementInfos.ExistsFills && !_context.GenerateElementInfos.ExistsFills && AllElements["fills"] < elementIndex)
        {
            await GenerateFillsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsFills = true;
        }
        else if (!_context.OldElementInfos.ExistsBorders && !_context.GenerateElementInfos.ExistsBorders && AllElements["borders"] < elementIndex)
        {
            await GenerateBordersAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsBorders = true;
        }
        else if (!_context.OldElementInfos.ExistsCellStyleXfs && !_context.GenerateElementInfos.ExistsCellStyleXfs && AllElements["cellStyleXfs"] < elementIndex)
        {
            await GenerateCellStyleXfsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsCellStyleXfs = true;
        }
        else if (!_context.OldElementInfos.ExistsCellXfs && !_context.GenerateElementInfos.ExistsCellXfs && AllElements["cellXfs"] < elementIndex)
        {
            await GenerateCellXfsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsCellXfs = true;
        }
    }

    [CreateSyncVersion]
    protected virtual async Task GenerateElementBeforEndElementAsync()
    {
        switch (OldReader.LocalName)
        {
            case "styleSheet" when !_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts:
                await GenerateNumFmtsAsync().ConfigureAwait(false);
                break;
            case "numFmts":
                await GenerateNumFmtAsync().ConfigureAwait(false);
                break;
            case "fonts":
                await GenerateFontAsync().ConfigureAwait(false);
                break;
            case "fills":
                await GenerateFillAsync().ConfigureAwait(false);
                break;
            case "borders":
                await GenerateBorderAsync().ConfigureAwait(false);
                break;
            case "cellStyleXfs":
                await GenerateCellStyleXfAsync().ConfigureAwait(false);
                break;
            case "cellXfs":
                await GenerateCellXfAsync().ConfigureAwait(false);
                break;
        }
    }

    [CreateSyncVersion]
    protected virtual async Task GenerateNumFmtsAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "numFmts", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateNumFmtAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsFonts)
        {
            await GenerateFontsAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected abstract Task GenerateNumFmtAsync();

    [CreateSyncVersion]
    protected virtual async Task GenerateFontsAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fonts", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString()).ConfigureAwait(false);
        await GenerateFontAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsFills)
        {
            await GenerateFillsAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected abstract Task GenerateFontAsync();

    [CreateSyncVersion]
    protected virtual async Task GenerateFillsAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fills", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString()).ConfigureAwait(false);
        await GenerateFillAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsBorders)
        {
            await GenerateBordersAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected abstract Task GenerateFillAsync();

    [CreateSyncVersion]
    protected virtual async Task GenerateBordersAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "borders", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString()).ConfigureAwait(false);
        await GenerateBorderAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsCellStyleXfs)
        {
            await GenerateCellStyleXfsAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected abstract Task GenerateBorderAsync();

    [CreateSyncVersion]
    protected virtual async Task GenerateCellStyleXfsAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "cellStyleXfs", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString()).ConfigureAwait(false);
        await GenerateCellStyleXfAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsCellXfs)
        {
            await GenerateCellXfsAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected abstract Task GenerateCellStyleXfAsync();

    [CreateSyncVersion]
    protected virtual async Task GenerateCellXfsAsync()
    {
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "cellXfs", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateCellXfAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected abstract Task GenerateCellXfAsync();

    private Dictionary<string, string> GetCellXfIdMap()
    {
        var result = new Dictionary<string, string>();
        for (int i = 0; i < _context.GenerateElementInfos.CellXfCount; i++)
        {
            result.Add(i.ToString(), (_context.OldElementInfos.CellXfCount + i).ToString());
        }
        return result;
    }
}