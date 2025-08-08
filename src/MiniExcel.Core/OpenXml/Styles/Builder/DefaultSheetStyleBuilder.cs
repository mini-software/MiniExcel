using System.Drawing;
using MiniExcelLib.Core.Enums;

namespace MiniExcelLib.Core.OpenXml.Styles.Builder;

internal partial class DefaultSheetStyleBuilder(SheetStyleBuildContext context, OpenXmlStyleOptions styleOptions)
    : SheetStyleBuilderBase(context)
{
    private static readonly SheetStyleElementInfos GenerateElementInfos = new()
    {
        NumFmtCount = 0, //The default NumFmt number is 0, but there will be NumFmt dynamically generated based on ColumnsToApply
        FontCount = 2,
        FillCount = 3,
        BorderCount = 2,
        CellStyleXfCount = 3,
        CellXfCount = 5
    };

    private static readonly Color DefaultBackgroundColor = Color.FromArgb(0x284472C4);
    private const HorizontalCellAlignment DefaultHorizontalAlignment = HorizontalCellAlignment.Left;
    private const VerticalCellAlignment DefaultVerticalAlignment = VerticalCellAlignment.Bottom;
    
    private readonly SheetStyleBuildContext _context = context;
    private readonly OpenXmlStyleOptions _styleOptions = styleOptions;

    private XmlReader OldReader => _context.OldXmlReader!;
    private XmlWriter NewWriter => _context.NewXmlWriter!; 

    protected override SheetStyleElementInfos GetGenerateElementInfos()
    {
        return GenerateElementInfos;
    }

    [CreateSyncVersion]
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
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "numFmt", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "formatCode", null, item.Format).ConfigureAwait(false);
            await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
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
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "font", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "vertAlign", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "baseline").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "sz", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "11").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "name", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "Calibri").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "family", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "2").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:font>
         *     <x:vertAlign val="baseline" />
         *     <x:sz val="11" />
         *     <x:color rgb="FFFFFFFF" />
         *     <x:name val="Calibri" />
         *     <x:family val="2" />
         * </x:font>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "font", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "vertAlign", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "baseline").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "sz", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "11").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FFFFFFFF").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "name", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "Calibri").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "family", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "val", null, "2").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill>
         *     <x:patternFill patternType="none" />
         * </x:fill>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "patternFill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "patternType", null, "none").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:fill>
         *     <x:patternFill patternType="gray125" />
         * </x:fill>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "patternFill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "patternType", null, "gray125").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:fill>
         *     <x:patternFill patternType="solid">
         *         <x:fgColor rgb="284472C4" />
         *     </x:patternFill>
         * </x:fill>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "patternFill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "patternType", null, "solid").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fgColor", OldReader.NamespaceURI).ConfigureAwait(false);

        var bgColor = _styleOptions.HeaderStyle?.BackgroundColor ?? DefaultBackgroundColor;
        var hexBgColor = $"{bgColor.A:X2}{bgColor.R:X2}{bgColor.G:X2}{bgColor.B:X2}";
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, hexBgColor).ConfigureAwait(false);
        
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
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
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "border", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "left", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "right", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "top", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "bottom", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "diagonal", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

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
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "border", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "diagonalUp", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "diagonalDown", null, "0").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "left", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "right", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "top", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "bottom", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "thin").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "diagonal", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "style", null, "none").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "color", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "rgb", null, "FF000000").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyNumberFormat="1" applyFill="1" applyBorder="0" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="14" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateCellXfAsync()
    {
        /*
         * <x:xf></x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFill="0" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="left" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 2}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "alignment", OldReader.NamespaceURI).ConfigureAwait(false);
        
        var horizontalAlignment = _styleOptions.HeaderStyle?.HorizontalAlignment ?? DefaultHorizontalAlignment;
        var horizontalAlignmentStr = horizontalAlignment.ToString().ToLowerInvariant();
        await NewWriter.WriteAttributeStringAsync(null, "horizontal", null, horizontalAlignmentStr).ConfigureAwait(false);
        
        var verticalAlignment = _styleOptions.HeaderStyle?.VerticalAlignment ?? DefaultVerticalAlignment;
        var verticalAlignmentStr = verticalAlignment.ToString().ToLowerInvariant();
        await NewWriter.WriteAttributeStringAsync(null, "vertical", null, verticalAlignmentStr).ConfigureAwait(false);
        
        await NewWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);

        var wrapHeader = (_styleOptions.HeaderStyle?.WrapText ?? false) ? "1" : "0";
        await NewWriter.WriteAttributeStringAsync(null, "wrapText", null, wrapHeader).ConfigureAwait(false);
        
        await NewWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "alignment", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
        
        var wrapContent = _styleOptions.WrapCellContents ? "1" : "0";
        await NewWriter.WriteAttributeStringAsync(null, "wrapText", null, wrapContent).ConfigureAwait(false);
        
        await NewWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="14" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
         *     <x:alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="0" indent="0" relativeIndent="0" justifyLastLine="0" shrinkToFit="0" readingOrder="0" />
         *     <x:protection locked="1" hidden="0" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "alignment", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "wrapText", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        /*
         * <x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
         *     <x:alignment horizontal="fill" />
         * </x:xf>
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "alignment", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "horizontal", null, "fill").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        var index = 0;
        foreach (var _ in _context.ColumnsToApply)
        {
            index++;

            /*
             * <x:xf numFmtId=""{numFmtIndex + i}"" fontId=""0"" fillId=""0"" borderId=""1"" xfId=""0"" applyNumberFormat=""1"" applyFill=""1"" applyBorder=""1"" applyAlignment=""1"" applyProtection=""1"">
             *     <x:alignment horizontal=""general"" vertical=""bottom"" textRotation=""0"" wrapText=""0"" indent=""0"" relativeIndent=""0"" justifyLastLine=""0"" shrinkToFit=""0"" readingOrder=""0"" />
             *     <x:protection locked=""1"" hidden=""0"" />
             * </x:xf>
             */
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fontId", null, $"{_context.OldElementInfos.FontCount + 0}").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fillId", null, $"{_context.OldElementInfos.FillCount + 0}").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "borderId", null, $"{_context.OldElementInfos.BorderCount + 1}").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyFill", null, "1").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyBorder", null, "1").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyAlignment", null, "1").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyProtection", null, "1").ConfigureAwait(false);
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "alignment", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "horizontal", null, "general").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "vertical", null, "bottom").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "textRotation", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "wrapText", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "indent", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "relativeIndent", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "justifyLastLine", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "shrinkToFit", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "readingOrder", null, "0").ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "protection", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "locked", null, "1").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "hidden", null, "0").ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        }
    }
}