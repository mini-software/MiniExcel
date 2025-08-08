namespace MiniExcelLib.Core.OpenXml.Styles.Builder;

internal partial class MinimalSheetStyleBuilder(SheetStyleBuildContext context) : SheetStyleBuilderBase(context)
{
    internal static SheetStyleElementInfos GenerateElementInfos = new()
    {
        NumFmtCount = 0, //The default NumFmt number is 0, but there will be NumFmt dynamically generated based on ColumnsToApply
        FontCount = 1,
        FillCount = 1,
        BorderCount = 1,
        CellStyleXfCount = 1,
        CellXfCount = 5
    };

    private readonly SheetStyleBuildContext _context = context;
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
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
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
         * <x:font />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "font", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "fill", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateBorderAsync()
    {
        /*
         * <x:border />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "border", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [CreateSyncVersion]
    protected override async Task GenerateCellXfAsync()
    {
        /*
         * <x:xf />
         * <x:xf />
         * <x:xf />
         * <x:xf numFmtId="14" applyNumberFormat="1" />
         * <x:xf />
         */
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
        await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        var index = 0;
        foreach (var _ in _context.ColumnsToApply)
        {
            index++;

            /*
             * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1"
             */
            await NewWriter.WriteStartElementAsync(OldReader.Prefix, "xf", OldReader.NamespaceURI).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index).ToString()).ConfigureAwait(false);
            await NewWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
            await NewWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        }
    }
}