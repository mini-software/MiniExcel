namespace MiniExcelLib.Core.OpenXml.Styles;

internal partial class MinimalSheetStyleBuilder : SheetStyleBuilderBase
{
    internal static SheetStyleElementInfos GenerateElementInfos = new SheetStyleElementInfos
    {
        NumFmtCount = 0,//默认的NumFmt数量是0，但是会有根据ColumnsToApply动态生成的NumFmt
        FontCount = 1,
        FillCount = 1,
        BorderCount = 1,
        CellStyleXfCount = 1,
        CellXfCount = 5
    };

    private readonly SheetStyleBuildContext _context;

    public MinimalSheetStyleBuilder(SheetStyleBuildContext context) : base(context)
    {
        _context = context;
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
             * <x:numFmt numFmtId="{numFmtIndex + i}" formatCode="{item.Format}" />
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "numFmt", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index + _context.OldElementInfos.NumFmtCount).ToString()).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "formatCode", null, item.Format).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        }
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateFontAsync()
    {
        /*
         * <x:font />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "font", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateFillAsync()
    {
        /*
         * <x:fill />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "fill", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateBorderAsync()
    {
        /*
         * <x:border />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "border", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateCellStyleXfAsync()
    {
        /*
         * <x:xf />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
    }

    [Zomp.SyncMethodGenerator.CreateSyncVersion]
    protected override async Task GenerateCellXfAsync()
    {
        /*
         * <x:xf />
         * <x:xf />
         * <x:xf />
         * <x:xf numFmtId="14" applyNumberFormat="1" />
         * <x:xf />
         */
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, "14").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
        await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);

        const int numFmtIndex = 166;
        var index = 0;
        foreach (var item in _context.ColumnsToApply)
        {
            index++;

            /*
             * <x:xf numFmtId="{numFmtIndex + i}" applyNumberFormat="1"
             */
            await _context.NewXmlWriter.WriteStartElementAsync(_context.OldXmlReader.Prefix, "xf", _context.OldXmlReader.NamespaceURI).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "numFmtId", null, (numFmtIndex + index).ToString()).ConfigureAwait(false);
            await _context.NewXmlWriter.WriteAttributeStringAsync(null, "applyNumberFormat", null, "1").ConfigureAwait(false);
            await _context.NewXmlWriter.WriteFullEndElementAsync().ConfigureAwait(false);
        }
    }
}