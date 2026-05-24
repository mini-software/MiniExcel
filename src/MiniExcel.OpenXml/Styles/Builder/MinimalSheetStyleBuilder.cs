namespace MiniExcelLib.OpenXml.Styles.Builder;

internal partial class MinimalSheetStyleBuilder(SheetStyleBuilderContext context) : SheetStyleBuilderBase(context)
{
    private static readonly SheetStyleElementInfos GenerateElementInfos = new()
    {
        NumFmtCount = 0, //The default NumFmt number is 0, but others will be dynamically generated based on format mappings
        FontCount = 1,
        FillCount = 1,
        BorderCount = 1,
        CellStyleXfCount = 1,
        CellXfCount = 6
    };
    
    protected internal override SheetStyleElementInfos GetGeneratedElementInfos()
    {
        return GenerateElementInfos;
    }

    [CreateSyncVersion]
    protected override async Task GenerateNumFmtAsync()
    {
        const int numFmtIndex = 166;
        var index = 0;
        foreach (var map in Context.SheetStyleFormatsCache.FormatMappings)
        {
            index++;

            /*
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
             */
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "numFmt", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + Context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "formatCode", null, map.Format).ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        }
    }

    [CreateSyncVersion]
    protected override async Task GenerateFontAsync()
    {
        /*
         * <x:font />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "font", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateBorderAsync()
    {
        /*
         * <x:border />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "border", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
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
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        }

        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "21").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        for (var i = 1; i <= Context.CustomFormatCount; i++)
        {
            /*
             * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1"
             */
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + i).ToString()).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fontId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "fillId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "borderId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "xfId", null, "0").ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
            await NewWriter.WriteEndElementAsync().ConfigureAwait(false);
        }
    }
}
