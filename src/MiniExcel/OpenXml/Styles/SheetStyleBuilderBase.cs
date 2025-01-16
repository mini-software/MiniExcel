using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml.Styles
{
    internal abstract class SheetStyleBuilderBase : ISheetStyleBuilder
    {
        internal readonly static Dictionary<string, int> _allElements = new Dictionary<string, int>
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

        public virtual SheetStyleBuildResult Build()
        {
            _context.Initialize(GetGenerateElementInfos());

            while (_context.OldXmlReader.Read())
            {
                switch (_context.OldXmlReader.NodeType)
                {
                    case XmlNodeType.Element:
                        GenerateElementBeforStartElement();
                        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, _context.OldXmlReader.LocalName, _context.OldXmlReader.NamespaceURI);
                        WriteAttributes(_context.OldXmlReader.LocalName);
                        if (_context.OldXmlReader.IsEmptyElement)
                        {
                            GenerateElementBeforEndElement();
                            _context.NewXmlWriter.WriteEndElement();
                        }
                        break;

                    case XmlNodeType.Text:
                        _context.NewXmlWriter.WriteString(_context.OldXmlReader.Value);
                        break;

                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        _context.NewXmlWriter.WriteWhitespace(_context.OldXmlReader.Value);
                        break;

                    case XmlNodeType.CDATA:
                        _context.NewXmlWriter.WriteCData(_context.OldXmlReader.Value);
                        break;

                    case XmlNodeType.EntityReference:
                        _context.NewXmlWriter.WriteEntityRef(_context.OldXmlReader.Name);
                        break;

                    case XmlNodeType.XmlDeclaration:
                    case XmlNodeType.ProcessingInstruction:
                        _context.NewXmlWriter.WriteProcessingInstruction(_context.OldXmlReader.Name, _context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.DocumentType:
                        _context.NewXmlWriter.WriteDocType(_context.OldXmlReader.Name, _context.OldXmlReader.GetAttribute("PUBLIC"), _context.OldXmlReader.GetAttribute("SYSTEM"), _context.OldXmlReader.Value);
                        break;

                    case XmlNodeType.Comment:
                        _context.NewXmlWriter.WriteComment(_context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.EndElement:
                        GenerateElementBeforEndElement();
                        _context.NewXmlWriter.WriteFullEndElement();
                        break;
                }
            }

            _context.FinalizeAndUpdateZipDictionary();

            return new SheetStyleBuildResult(GetCellXfIdMap());
        }

        public virtual async Task<SheetStyleBuildResult> BuildAsync(CancellationToken cancellationToken = default)
        {
            await _context.InitializeAsync(GetGenerateElementInfos());

            while (await _context.OldXmlReader.ReadAsync())
            {
                cancellationToken.ThrowIfCancellationRequested();

                switch (_context.OldXmlReader.NodeType)
                {
                    case XmlNodeType.Element:
                        await GenerateElementBeforStartElementAsync();
                        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, _context.OldXmlReader.LocalName, _context.OldXmlReader.NamespaceURI);
                        await WriteAttributesAsync(_context.OldXmlReader.LocalName);
                        if (_context.OldXmlReader.IsEmptyElement)
                        {
                            GenerateElementBeforEndElement();
                            await _context.NewXmlWriter.WriteEndElementAsync();
                        }
                        break;
                    case XmlNodeType.Text:
                        await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        await _context.NewXmlWriter.WriteWhitespaceAsync(_context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.CDATA:
                        await _context.NewXmlWriter.WriteCDataAsync(_context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.EntityReference:
                        await _context.NewXmlWriter.WriteEntityRefAsync(_context.OldXmlReader.Name);
                        break;
                    case XmlNodeType.XmlDeclaration:
                    case XmlNodeType.ProcessingInstruction:
                        await _context.NewXmlWriter.WriteProcessingInstructionAsync(_context.OldXmlReader.Name, _context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.DocumentType:
                        await _context.NewXmlWriter.WriteDocTypeAsync(_context.OldXmlReader.Name, _context.OldXmlReader.GetAttribute("PUBLIC"), _context.OldXmlReader.GetAttribute("SYSTEM"), _context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.Comment:
                        await _context.NewXmlWriter.WriteCommentAsync(_context.OldXmlReader.Value);
                        break;
                    case XmlNodeType.EndElement:
                        GenerateElementBeforEndElement();
                        await _context.NewXmlWriter.WriteFullEndElementAsync();
                        break;
                }
            }

            await _context.FinalizeAndUpdateZipDictionaryAsync();

            return new SheetStyleBuildResult(GetCellXfIdMap());
        }

        protected abstract SheetStyleElementInfos GetGenerateElementInfos();

        protected virtual void WriteAttributes(string element)
        {
            if (_context.OldXmlReader.NodeType is XmlNodeType.Element || _context.OldXmlReader.NodeType is XmlNodeType.XmlDeclaration)
            {
                if (_context.OldXmlReader.MoveToFirstAttribute())
                {
                    WriteAttributes(element);
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
                        if (_context.OldXmlReader.NodeType == XmlNodeType.EntityReference)
                        {
                            _context.NewXmlWriter.WriteEntityRef(_context.OldXmlReader.Name);
                        }
                        else if (currentAttribute == "count")
                        {
                            switch (element)
                            {
                                case "numFmts":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString());
                                    break;
                                case "fonts":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString());
                                    break;
                                case "fills":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString());
                                    break;
                                case "borders":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString());
                                    break;
                                case "cellStyleXfs":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString());
                                    break;
                                case "cellXfs":
                                    _context.NewXmlWriter.WriteString((_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString());
                                    break;
                                default:
                                    _context.NewXmlWriter.WriteString(_context.OldXmlReader.Value);
                                    break;
                            }
                        }
                        else
                        {
                            _context.NewXmlWriter.WriteString(_context.OldXmlReader.Value);
                        }
                    }
                    _context.NewXmlWriter.WriteEndAttribute();
                }
                while (_context.OldXmlReader.MoveToNextAttribute());
            }
        }

        protected virtual async Task WriteAttributesAsync(string element)
        {
            if (_context.OldXmlReader.NodeType is XmlNodeType.Element || _context.OldXmlReader.NodeType is XmlNodeType.XmlDeclaration)
            {
                if (_context.OldXmlReader.MoveToFirstAttribute())
                {
                    await WriteAttributesAsync(element);
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
                        if (_context.OldXmlReader.NodeType == XmlNodeType.EntityReference)
                        {
                            await _context.NewXmlWriter.WriteEntityRefAsync(_context.OldXmlReader.Name);
                        }
                        else if (currentAttribute == "count")
                        {
                            switch (element)
                            {
                                case "numFmts":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString());
                                    break;
                                case "fonts":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString());
                                    break;
                                case "fills":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString());
                                    break;
                                case "borders":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString());
                                    break;
                                case "cellStyleXfs":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString());
                                    break;
                                case "cellXfs":
                                    await _context.NewXmlWriter.WriteStringAsync((_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString());
                                    break;
                                default:
                                    await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value);
                                    break;
                            }
                        }
                        else
                        {
                            await _context.NewXmlWriter.WriteStringAsync(_context.OldXmlReader.Value);
                        }
                    }
                    _context.NewXmlWriter.WriteEndAttribute();
                }
                while (_context.OldXmlReader.MoveToNextAttribute());
            }
        }

        protected virtual void GenerateElementBeforStartElement()
        {
            if (!_allElements.TryGetValue(_context.OldXmlReader.LocalName, out var elementIndex))
            {
                return;
            }
            if (!_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts && _allElements["numFmts"] < elementIndex)
            {
                GenerateNumFmts();
                _context.GenerateElementInfos.ExistsNumFmts = true;
            }
            else if (!_context.OldElementInfos.ExistsFonts && !_context.GenerateElementInfos.ExistsFonts && _allElements["fonts"] < elementIndex)
            {
                GenerateFonts();
                _context.GenerateElementInfos.ExistsFonts = true;
            }
            else if (!_context.OldElementInfos.ExistsFills && !_context.GenerateElementInfos.ExistsFills && _allElements["fills"] < elementIndex)
            {
                GenerateFills();
                _context.GenerateElementInfos.ExistsFills = true;
            }
            else if (!_context.OldElementInfos.ExistsBorders && !_context.GenerateElementInfos.ExistsBorders && _allElements["borders"] < elementIndex)
            {
                GenerateBorders();
                _context.GenerateElementInfos.ExistsBorders = true;
            }
            else if (!_context.OldElementInfos.ExistsCellStyleXfs && !_context.GenerateElementInfos.ExistsCellStyleXfs && _allElements["cellStyleXfs"] < elementIndex)
            {
                GenerateCellStyleXfs();
                _context.GenerateElementInfos.ExistsCellStyleXfs = true;
            }
            else if (!_context.OldElementInfos.ExistsCellXfs && !_context.GenerateElementInfos.ExistsCellXfs && _allElements["cellXfs"] < elementIndex)
            {
                GenerateCellXfs();
                _context.GenerateElementInfos.ExistsCellXfs = true;
            }
        }

        protected virtual async Task GenerateElementBeforStartElementAsync()
        {
            if (!_allElements.TryGetValue(_context.OldXmlReader.LocalName, out var elementIndex))
            {
                return;
            }
            if (!_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts && _allElements["numFmts"] < elementIndex)
            {
                await GenerateNumFmtsAsync();
                _context.GenerateElementInfos.ExistsNumFmts = true;
            }
            else if (!_context.OldElementInfos.ExistsFonts && !_context.GenerateElementInfos.ExistsFonts && _allElements["fonts"] < elementIndex)
            {
                await GenerateFontsAsync();
                _context.GenerateElementInfos.ExistsFonts = true;
            }
            else if (!_context.OldElementInfos.ExistsFills && !_context.GenerateElementInfos.ExistsFills && _allElements["fills"] < elementIndex)
            {
                await GenerateFillsAsync();
                _context.GenerateElementInfos.ExistsFills = true;
            }
            else if (!_context.OldElementInfos.ExistsBorders && !_context.GenerateElementInfos.ExistsBorders && _allElements["borders"] < elementIndex)
            {
                await GenerateBordersAsync();
                _context.GenerateElementInfos.ExistsBorders = true;
            }
            else if (!_context.OldElementInfos.ExistsCellStyleXfs && !_context.GenerateElementInfos.ExistsCellStyleXfs && _allElements["cellStyleXfs"] < elementIndex)
            {
                await GenerateCellStyleXfsAsync();
                _context.GenerateElementInfos.ExistsCellStyleXfs = true;
            }
            else if (!_context.OldElementInfos.ExistsCellXfs && !_context.GenerateElementInfos.ExistsCellXfs && _allElements["cellXfs"] < elementIndex)
            {
                await GenerateCellXfsAsync();
                _context.GenerateElementInfos.ExistsCellXfs = true;
            }
        }

        protected virtual void GenerateElementBeforEndElement()
        {
            if (_context.OldXmlReader.LocalName == "styleSheet" && !_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts)
            {
                GenerateNumFmts();
            }
            else if (_context.OldXmlReader.LocalName == "numFmts")
            {
                GenerateNumFmt();
            }
            else if (_context.OldXmlReader.LocalName == "fonts")
            {
                GenerateFont();
            }
            else if (_context.OldXmlReader.LocalName == "fills")
            {
                GenerateFill();
            }
            else if (_context.OldXmlReader.LocalName == "borders")
            {
                GenerateBorder();
            }
            else if (_context.OldXmlReader.LocalName == "cellStyleXfs")
            {
                GenerateCellStyleXf();
            }
            else if (_context.OldXmlReader.LocalName == "cellXfs")
            {
                GenerateCellXf();
            }
        }

        protected virtual async Task GenerateElementBeforEndElementAsync()
        {
            if (_context.OldXmlReader.LocalName == "styleSheet" && !_context.OldElementInfos.ExistsNumFmts && !_context.GenerateElementInfos.ExistsNumFmts)
            {
                await GenerateNumFmtsAsync();
            }
            else if (_context.OldXmlReader.LocalName == "numFmts")
            {
                await GenerateNumFmtAsync();
            }
            else if (_context.OldXmlReader.LocalName == "fonts")
            {
                await GenerateFontAsync();
            }
            else if (_context.OldXmlReader.LocalName == "fills")
            {
                await GenerateFillAsync();
            }
            else if (_context.OldXmlReader.LocalName == "borders")
            {
                await GenerateBorderAsync();
            }
            else if (_context.OldXmlReader.LocalName == "cellStyleXfs")
            {
                await GenerateCellStyleXfAsync();
            }
            else if (_context.OldXmlReader.LocalName == "cellXfs")
            {
                await GenerateCellXfAsync();
            }
        }

        protected virtual void GenerateNumFmts()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "numFmts", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString());
            GenerateNumFmt();
            _context.NewXmlWriter.WriteFullEndElement();

            if (!_context.OldElementInfos.ExistsFonts)
            {
                GenerateFonts();
            }
        }

        protected virtual async Task GenerateNumFmtsAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmts", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.NumFmtCount + _context.GenerateElementInfos.NumFmtCount + _context.CustomFormatCount).ToString());
            await GenerateNumFmtAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            if (!_context.OldElementInfos.ExistsFonts)
            {
                await GenerateFontsAsync();
            }
        }

        protected abstract void GenerateNumFmt();

        protected abstract Task GenerateNumFmtAsync();

        protected virtual void GenerateFonts()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fonts", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString());
            GenerateFont();
            _context.NewXmlWriter.WriteFullEndElement();

            if (!_context.OldElementInfos.ExistsFills)
            {
                GenerateFills();
            }
        }

        protected virtual async Task GenerateFontsAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fonts", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.FontCount + _context.GenerateElementInfos.FontCount).ToString());
            await GenerateFontAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            if (!_context.OldElementInfos.ExistsFills)
            {
                await GenerateFillsAsync();
            }
        }

        protected abstract void GenerateFont();

        protected abstract Task GenerateFontAsync();

        protected virtual void GenerateFills()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fills", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString());
            GenerateFill();
            _context.NewXmlWriter.WriteFullEndElement();
            if (!_context.OldElementInfos.ExistsBorders)
            {
                GenerateBorders();
            }
        }

        protected virtual async Task GenerateFillsAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fills", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.FillCount + _context.GenerateElementInfos.FillCount).ToString());
            await GenerateFillAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            if (!_context.OldElementInfos.ExistsBorders)
            {
                await GenerateBordersAsync();
            }
        }

        protected abstract void GenerateFill();

        protected abstract Task GenerateFillAsync();

        protected virtual void GenerateBorders()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "borders", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString());
            GenerateBorder();
            _context.NewXmlWriter.WriteFullEndElement();

            if (!_context.OldElementInfos.ExistsCellStyleXfs)
            {
                GenerateCellStyleXfs();
            }
        }

        protected virtual async Task GenerateBordersAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "borders", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.BorderCount + _context.GenerateElementInfos.BorderCount).ToString());
            await GenerateBorderAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            if (!_context.OldElementInfos.ExistsCellStyleXfs)
            {
                await GenerateCellStyleXfsAsync();
            }
        }

        protected abstract void GenerateBorder();

        protected abstract Task GenerateBorderAsync();

        protected virtual void GenerateCellStyleXfs()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "cellStyleXfs", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString());
            GenerateCellStyleXf();
            _context.NewXmlWriter.WriteFullEndElement();

            if (!_context.OldElementInfos.ExistsCellXfs)
            {
                GenerateCellXfs();
            }
        }

        protected virtual async Task GenerateCellStyleXfsAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "cellStyleXfs", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.CellStyleXfCount + _context.GenerateElementInfos.CellStyleXfCount).ToString());
            await GenerateCellStyleXfAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();

            if (!_context.OldElementInfos.ExistsCellXfs)
            {
                await GenerateCellXfsAsync();
            }
        }

        protected abstract void GenerateCellStyleXf();

        protected abstract Task GenerateCellStyleXfAsync();

        protected virtual void GenerateCellXfs()
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "cellXfs", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("count", (_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString());
            GenerateCellXf();
            _context.NewXmlWriter.WriteFullEndElement();
        }

        protected virtual async Task GenerateCellXfsAsync()
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "cellXfs", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(_context.OldXmlReader.Prefix, "count", _context.OldXmlReader.NamespaceURI, (_context.OldElementInfos.CellXfCount + _context.GenerateElementInfos.CellXfCount + _context.CustomFormatCount).ToString());
            await GenerateCellXfAsync();
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }

        protected abstract void GenerateCellXf();

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
}
