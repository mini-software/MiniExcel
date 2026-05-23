namespace MiniExcelLib.OpenXml.Styles.Builder;

internal abstract partial class SheetStyleBuilderBase(SheetStyleBuilderContext context) : ISheetStyleBuilder
{
    protected readonly SheetStyleBuilderContext Context = context;
    
    //todo: these may actually be null if used when the context is not initialized
    protected XmlReader OldReader => Context.OldXmlReader!;
    protected XmlWriter NewWriter => Context.NewXmlWriter!;

    private static readonly Dictionary<string, int> AllElements = new()
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
    public virtual async Task BuildAsync(CancellationToken cancellationToken = default)
    {
        await Context.InitializeAsync(GetGeneratedElementInfos(), cancellationToken).ConfigureAwait(false);
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

        await Context.FinalizeAndUpdateZipDictionaryAsync(cancellationToken).ConfigureAwait(false);
    }

    protected internal abstract SheetStyleElementInfos GetGeneratedElementInfos();

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
                            "numFmts" => (Context.OldElementInfos.NumFmtCount + Context.GeneratedElementInfos.NumFmtCount + Context.CustomFormatCount).ToString(),
                            "fonts" => (Context.OldElementInfos.FontCount + Context.GeneratedElementInfos.FontCount).ToString(),
                            "fills" => (Context.OldElementInfos.FillCount + Context.GeneratedElementInfos.FillCount).ToString(),
                            "borders" => (Context.OldElementInfos.BorderCount + Context.GeneratedElementInfos.BorderCount).ToString(),
                            "cellStyleXfs" => (Context.OldElementInfos.CellStyleXfCount + Context.GeneratedElementInfos.CellStyleXfCount).ToString(),
                            "cellXfs" => (Context.OldElementInfos.CellXfCount + Context.GeneratedElementInfos.CellXfCount + Context.CustomFormatCount).ToString(),
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
        
        if (!Context.OldElementInfos.ExistsNumFmts && !Context.GeneratedElementInfos.ExistsNumFmts && AllElements["numFmts"] < elementIndex)
        {
            await GenerateNumFmtsAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsNumFmts = true;
        }
        else if (!Context.OldElementInfos.ExistsFonts && !Context.GeneratedElementInfos.ExistsFonts && AllElements["fonts"] < elementIndex)
        {
            await GenerateFontsAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsFonts = true;
        }
        else if (!Context.OldElementInfos.ExistsFills && !Context.GeneratedElementInfos.ExistsFills && AllElements["fills"] < elementIndex)
        {
            await GenerateFillsAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsFills = true;
        }
        else if (!Context.OldElementInfos.ExistsBorders && !Context.GeneratedElementInfos.ExistsBorders && AllElements["borders"] < elementIndex)
        {
            await GenerateBordersAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsBorders = true;
        }
        else if (!Context.OldElementInfos.ExistsCellStyleXfs && !Context.GeneratedElementInfos.ExistsCellStyleXfs && AllElements["cellStyleXfs"] < elementIndex)
        {
            await GenerateCellStyleXfsAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsCellStyleXfs = true;
        }
        else if (!Context.OldElementInfos.ExistsCellXfs && !Context.GeneratedElementInfos.ExistsCellXfs && AllElements["cellXfs"] < elementIndex)
        {
            await GenerateCellXfsAsync().ConfigureAwait(false);
            Context.GeneratedElementInfos.ExistsCellXfs = true;
        }
    }

    [CreateSyncVersion]
    protected virtual async Task GenerateElementBeforEndElementAsync()
    {
        switch (OldReader.LocalName)
        {
            case "styleSheet" when !Context.OldElementInfos.ExistsNumFmts && !Context.GeneratedElementInfos.ExistsNumFmts:
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.NumFmtCount + Context.GeneratedElementInfos.NumFmtCount + Context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateNumFmtAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!Context.OldElementInfos.ExistsFonts)
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.FontCount + Context.GeneratedElementInfos.FontCount).ToString()).ConfigureAwait(false);
        await GenerateFontAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!Context.OldElementInfos.ExistsFills)
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.FillCount + Context.GeneratedElementInfos.FillCount).ToString()).ConfigureAwait(false);
        await GenerateFillAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!Context.OldElementInfos.ExistsBorders)
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.BorderCount + Context.GeneratedElementInfos.BorderCount).ToString()).ConfigureAwait(false);
        await GenerateBorderAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!Context.OldElementInfos.ExistsCellStyleXfs)
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.CellStyleXfCount + Context.GeneratedElementInfos.CellStyleXfCount).ToString()).ConfigureAwait(false);
        await GenerateCellStyleXfAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!Context.OldElementInfos.ExistsCellXfs)
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
        await NewWriter.WriteAttributeStringAsync(null, "count", null, (Context.OldElementInfos.CellXfCount + Context.GeneratedElementInfos.CellXfCount + Context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateCellXfAsync().ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected abstract Task GenerateCellXfAsync();
}