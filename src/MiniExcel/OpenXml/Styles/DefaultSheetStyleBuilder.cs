namespace MiniExcelLib.OpenXml.Styles;

internal partial class DefaultSheetStyleBuilder : SheetStyleBuilderBase
{
    private static readonly SheetStyleElementInfos GenerateElementInfos = new SheetStyleElementInfos
    {
        NumFmtCount = 0,//The default NumFmt number is 0, but there will be NumFmt dynamically generated based on ColumnsToApply
        FontCount = 2,
        FillCount = 3,
        BorderCount = 2,
        CellStyleXfCount = 3,
        CellXfCount = 5
    };

    private readonly SheetStyleBuildContext _context;
    private OpenXmlStyleOptions _styleOptions;

    public DefaultSheetStyleBuilder(SheetStyleBuildContext context, OpenXmlStyleOptions styleOptions) : base(context)
    {
        _context = context;
        _styleOptions = styleOptions;
    }

    protected override SheetStyleElementInfos GetGenerateElementInfos()
    {
        return GenerateElementInfos;
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateNumFmtAsync()
    {
        const int numFmtIndex = 166;
        var index = 0;
        foreach (var item in _context.ColumnsToApply)
        {
            index++;

            /*
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{x.Format}"
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false); ;
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "formatCode", null, item.Format).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateFontAsync()
    {
        /*
         * <x:font>
         *     <x:vertAlign val="baseline" />
         *     <x:sz val="11" />
         *     <x:color rgb="FF000000" />
         *     <x:name val="Calibri" />
         *     <x:family val="2" />
         * </x:font>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "baseline").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "11").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "Calibri").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "2").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:font>
         *     <x:vertAlign val="baseline" />
         *     <x:sz val="11" />
         *     <x:color rgb="FFFFFFFF" />
         *     <x:name val="Calibri" />
         *     <x:family val="2" />
         * </x:font>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "vertAlign", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "baseline").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "sz", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "11").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FFFFFFFF").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "name", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "Calibri").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "family", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "val", null, "2").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill>
         *     <x:patternFill patternType="none" />
         * </x:fill>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:fill>
         *     <x:patternFill patternType="gray125" />
         * </x:fill>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "gray125").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:fill>
         *     <x:patternFill patternType="solid">
         *         <x:fgColor rgb="284472C4" />
         *     </x:patternFill>
         * </x:fill>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "patternFill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "patternType", null, "solid").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fgColor", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "284472C4").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateBorderAsync()
    {
        /*
         * <x:border diagonalUp="0" diagonalDown="0">
         *     <x:left style="none">
         *         <x:color rgb="FF000000" />
         *     </x:left>
         *     <x:right style="none">
         *         <x:color rgb="FF000000" />
         *     </x:right>
         *     <x:top style="none">
         *         <x:color rgb="FF000000" />
         *     </x:top>
         *     <x:bottom style="none">
         *         <x:color rgb="FF000000" />
         *     </x:bottom>
         *     <x:diagonal style="none">
         *         <x:color rgb="FF000000" />
         *     </x:diagonal>
         * </x:border>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:border diagonalUp="0" diagonalDown="0">
         *     <x:left style="thin">
         *         <x:color rgb="FF000000" />
         *     </x:left>
         *     <x:right style="thin">
         *         <x:color rgb="FF000000" />
         *     </x:right>
         *     <x:top style="thin">
         *         <x:color rgb="FF000000" />
         *     </x:top>
         *     <x:bottom style="thin">
         *         <x:color rgb="FF000000" />
         *     </x:bottom>
         *     <x:diagonal style="none">
         *         <x:color rgb="FF000000" />
         *     </x:diagonal>
         * </x:border>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "left", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "right", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "top", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "bottom", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "diagonal", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "color", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFill="1" applyBorder="0" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="14" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateCellXfAsync()
    {
        /*
         * <x:xf></x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="left" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "left").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, _styleOptions.WrapCellContents ? "1" : "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
         *     <x:alignment horizontal="fill" />
         * </x:xf>
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "fill").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        var index = 0;
        foreach (var item in _context.ColumnsToApply)
        {
            index++;

            /*
             * <x:xf numFmtId=""{numFmtIndex + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
             *     <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
             *     <x:protection locked=""1"" hidden=""0"" />
             * </x:xf>
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "alignment", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "wrapText", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "protection", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
            await _context.NewXmlWriter.WriteEndElementAsync().ConfigureAwait(false);
        }
    }
}