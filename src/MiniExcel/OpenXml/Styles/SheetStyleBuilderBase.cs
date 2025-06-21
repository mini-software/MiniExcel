using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles;

internal abstract partial class SheetStyleBuilderBase : ISheetStyleBuilder
{
    internal static readonly Dictionary<string, int> _allElements = new Dictionary<string, int>
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

    private readonly SheetStyleBuildContext _context;

    public SheetStyleBuilderBase(SheetStyleBuildContext context)
    {
        _context = context;
    }

    // Todo: add CancellationToken to all methods called inside of BuildAsync
    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    public virtual async Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default)
    {
        await _context.InitializeAsync(GetGenerateElementInfos(), cancellationToken).ConfigureAwait(false);

        while (await _context.OldXmlReader.ReadAsync().ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();

            switch (_context.OldXmlReader.NodeType)
            {
                case XmlNodeType.Element:
                    await GenerateElementBeforStartElementAsync().ConfigureAwait(false);
                    await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, _context.OldXmlReader.LocalName, _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
                    await WriteAttributesAsync(_context.OldXmlReader.LocalName, cancellationToken).ConfigureAwait(false);
                    if (_context.OldXmlReader.IsEmptyElement)
                    {
                        await GenerateElementBeforEndElementAsync().ConfigureAwait(false);
                        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
                    }
                    break;
                case XmlNodeType.Text:
                    await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.Whitespace:
                case XmlNodeType.SignificantWhitespace:
                    await _context.NewXmlWriter.WriteWhitespaceAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.CDATA:
                    await _context.NewXmlWriter.WriteCDataAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.EntityReference:
                    await _context.NewXmlWriter.WriteEntityRefAsync(_context.OldXmlReader.Name).ConfigureAwait(false);
                    break;
                case XmlNodeType.XmlDeclaration:
                case XmlNodeType.ProcessingInstruction:
                    await _context.NewXmlWriter.WriteProcessingInstructionAsync(_context.OldXmlReader.Name, _context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.DocumentType:
                    await _context.NewXmlWriter.WriteDocTypeAsync(_context.OldXmlReader.Name, _context.OldXmlReader.GetAttribute("PUBLIC"), _context.OldXmlReader.GetAttribute("SYSTEM"), _context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.Comment:
                    await _context.NewXmlWriter.WriteCommentAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                    break;
                case XmlNodeType.EndElement:
                    await GenerateElementBeforEndElementAsync().ConfigureAwait(false);
                    await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
                    break;
            }
        }

        await _context.FinalizeAndUpdateZipDictionaryAsync(cancellationToken).ConfigureAwait(false);

        return new SheetStyleBuildResult(GetCellXfIdMap());
    }

    protected abstract SheetStyleElementInfos GetGenerateElementInfos();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task WriteAttributesAsync(string element, CancellationToken cancellationToken = default)
    {
        if (_context.OldXmlReader.NodeType is XmlNodeType.Element || _context.OldXmlReader.NodeType is XmlNodeType.XmlDeclaration)
        {
            if (_context.OldXmlReader.MoveToFirstAttribute())
            {
                await WriteAttributesAsync(element, cancellationToken).ConfigureAwait(false);
                _context.OldXmlReader.MoveToElement();
            }
        }
        else if (_context.OldXmlReader.NodeType == XmlNodeType.Attribute)
        {
            do
            {
                _context.NewXmlWriter.WriteStartAttribute(_context.OldXmlReader.Prefix, _context.OldXmlReader.LocalName, _context.OldXmlReader.NamespaceURI);
                var currentAttribute = _context.OldXmlReader.LocalName;
                while (_context.OldXmlReader.ReadAttributeValue())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                        
                    if (_context.OldXmlReader.NodeType == XmlNodeType.EntityReference)
                    {
                        await _context.NewXmlWriter.WriteEntityRefAsync(_context.OldXmlReader.Name).ConfigureAwait(false);
                    }
                    else if (currentAttribute == "count")
                    {
                        switch (element)
                        {
                            case "numFmts":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
                                break;
                            case "fonts":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString()).ConfigureAwait(false);
                                break;
                            case "fills":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString()).ConfigureAwait(false);
                                break;
                            case "borders":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString()).ConfigureAwait(false);
                                break;
                            case "cellStyleXfs":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString()).ConfigureAwait(false);
                                break;
                            case "cellXfs":
                                await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
                                break;
                            default:
                                await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                                break;
                        }
                    }
                    else
                    {
                        await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value).ConfigureAwait(false);
                    }
                }
                _context.NewXmlWriter.WriteEndAttribute();
            }
            while (_context.OldXmlReader.MoveToNextAttribute());
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateElementBeforStartElementAsync()
    {
        if (!_allElements.TryGetValue(_context.OldXmlReader.LocalName, out var elementIndex))
        {
            return;
        }
        if (!_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts && _allElements["numFmts"] < elementIndex)
        {
            await GenerateNumFmtsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsNumFmts = true;
        }
        else if (!_context.OldElementInfos.ExistsFonts && !_context.GenerateElementInfos.ExistsFonts && _allElements["fonts"] < elementIndex)
        {
            await GenerateFontsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsFonts = true;
        }
        else if (!_context.OldElementInfos.ExistsFills && !_context.GenerateElementInfos.ExistsFills && _allElements["fills"] < elementIndex)
        {
            await GenerateFillsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsFills = true;
        }
        else if (!_context.OldElementInfos.ExistsBorders && !_context.GenerateElementInfos.ExistsBorders && _allElements["borders"] < elementIndex)
        {
            await GenerateBordersAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsBorders = true;
        }
        else if (!_context.OldElementInfos.ExistsCellStyleXfs && !_context.GenerateElementInfos.ExistsCellStyleXfs && _allElements["cellStyleXfs"] < elementIndex)
        {
            await GenerateCellStyleXfsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsCellStyleXfs = true;
        }
        else if (!_context.OldElementInfos.ExistsCellXfs && !_context.GenerateElementInfos.ExistsCellXfs && _allElements["cellXfs"] < elementIndex)
        {
            await GenerateCellXfsAsync().ConfigureAwait(false);
            _context.GenerateElementInfos.ExistsCellXfs = true;
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateElementBeforEndElementAsync()
    {
        switch (_context.OldXmlReader.LocalName)
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

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateNumFmtsAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmts", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateNumFmtAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsFonts)
        {
            await GenerateFontsAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected abstract Task GenerateNumFmtAsync();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateFontsAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fonts", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString()).ConfigureAwait(false);
        await GenerateFontAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsFills)
        {
            await GenerateFillsAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected abstract Task GenerateFontAsync();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateFillsAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fills", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString()).ConfigureAwait(false);
        await GenerateFillAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsBorders)
        {
            await GenerateBordersAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected abstract Task GenerateFillAsync();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateBordersAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "borders", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString()).ConfigureAwait(false);
        await GenerateBorderAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsCellStyleXfs)
        {
            await GenerateCellStyleXfsAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected abstract Task GenerateBorderAsync();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateCellStyleXfsAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "cellStyleXfs", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString()).ConfigureAwait(false);
        await GenerateCellStyleXfAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        if (!_context.OldElementInfos.ExistsCellXfs)
        {
            await GenerateCellXfsAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected abstract Task GenerateCellStyleXfAsync();

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected virtual async Task GenerateCellXfsAsync()
    {
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "cellXfs", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "count", null, (_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString()).ConfigureAwait(false);
        await GenerateCellXfAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
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