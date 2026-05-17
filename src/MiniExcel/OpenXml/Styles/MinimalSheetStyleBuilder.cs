namespace MiniExcelLibs.OpenXml.Styles;

internal class MinimalSheetStyleBuilder(SheetStyleBuildContext context) : SheetStyleBuilderBase(context)
{
    private static readonly SheetStyleElementInfos GenerateElementInfos = new()
    {
        NumFmtCount = 0,
        FontCount = 1,
        FillCount = 1,
        BorderCount = 1,
        CellStyleXfCount = 1,
        CellXfCount = 6
    };

    private readonly SheetStyleBuildContext _context = context;

    protected internal override SheetStyleElementInfos GetGenerateElementInfos()
    {
        return GenerateElementInfos;
    }

    protected override void GenerateNumFmt()
    {
        const int numFmtIndex = 166;

        var index = 0;
        foreach (var item in _context.SheetStyleFormatsCache.FormatMappings)
        {
            index++;

            /*
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
            _context.NewXmlWriter.WriteAttributeString("formatCode", item.Key);
            _context.NewXmlWriter.WriteFullEndElement();
        }
    }

    protected override async Task GenerateNumFmtAsync()
    {
        const int numFmtIndex = 166;
        var index = 0;
        foreach (var item in _context.SheetStyleFormatsCache.FormatMappings)
        {
            index++;

            /*
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString());
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "formatCode", null, item.Key);
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }
    }

    protected override void GenerateFont()
    {
        /*
         * <x:font />
         */
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteFullEndElement();
    }

    protected override async Task GenerateFontAsync()
    {
        /*
         * <x:font />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI);
        await _context.NewXmlWriter.WriteFullEndElementAsync();
    }

    protected override void GenerateFill()
    {
        /*
         * <x:fill />
         */
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteFullEndElement();
    }

    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI);
        await _context.NewXmlWriter.WriteFullEndElementAsync();
    }

    protected override void GenerateBorder()
    {
        /*
         * <x:border />
         */
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteFullEndElement();
    }

    protected override async Task GenerateBorderAsync()
    {
        /*
         * <x:border />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI);
        await _context.NewXmlWriter.WriteFullEndElementAsync();
    }

    protected override void GenerateCellStyleXf()
    {
        /*
         * <x:xf />
         */
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteFullEndElement();
    }

    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
        await _context.NewXmlWriter.WriteFullEndElementAsync();
    }

    protected override void GenerateCellXf()
    {
        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="21" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
         * */

        for (int i = 0; i < 3; i++)
        {
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("fontId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("fillId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("borderId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("xfId", null, "0");
            _context.NewXmlWriter.WriteEndElement();
        }

        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteAttributeString("numFmtId", null, "14");
        _context.NewXmlWriter.WriteAttributeString("fontId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("fillId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("borderId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("xfId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", null, "1");
        _context.NewXmlWriter.WriteEndElement();
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteAttributeString("numFmtId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("fontId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("fillId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("borderId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("xfId", null, "0");
        _context.NewXmlWriter.WriteEndElement();
        _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
        _context.NewXmlWriter.WriteAttributeString("numFmtId", null, "21");
        _context.NewXmlWriter.WriteAttributeString("fontId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("fillId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("borderId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("xfId", null, "0");
        _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", null, "1");
        _context.NewXmlWriter.WriteEndElement();

        const int numFmtIndex = 166;
        for (var i = 1; i <= _context.CustomFormatCount; i++)
        {
            /*
             * <x:xf numFmtId="{numFmtIndex + i}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyNumberFormat="1"
             */
            _context.NewXmlWriter.WriteStartElement(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            _context.NewXmlWriter.WriteAttributeString("numFmtId", (numFmtIndex + i).ToString());
            _context.NewXmlWriter.WriteAttributeString("fontId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("fillId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("borderId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("xfId", null, "0");
            _context.NewXmlWriter.WriteAttributeString("applyNumberFormat", "1");
            _context.NewXmlWriter.WriteEndElement();
        }
    }

    protected override async Task GenerateCellXfAsync()
    {
        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
         * <x:xf numFmtId="21" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
         * */

        for (int i = 0; i < 3; i++)
        {
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        }

        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "21").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        for (var i = 1; i <= _context.CustomFormatCount; i++)
        {
            /*
             * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1"
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + i).ToString());
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1");
            await _context.NewXmlWriter.WriteFullEndElementAsync();
        }
    }
}