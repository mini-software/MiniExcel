namespace MiniExcelLib.Core.OpenXml.Styles;

public static class SheetStyleBuilderHelper
{
    public static IEnumerable<MiniExcelColumnAttribute> GenerateStyleIds(int startUpCellXfs, ICollection<MiniExcelColumnAttribute>? dynamicColumns)
    {
        if (dynamicColumns is null)
            yield break;

        int index = 0;
        var cols = dynamicColumns
            .Where(x => !string.IsNullOrWhiteSpace(x.Format) && new OpenXmlNumberFormatHelper(x.Format).IsValid)
            .GroupBy(x => x.Format);
        
        foreach (var g in cols) 
        {
            foreach ( var col in g )
                col.FormatId = startUpCellXfs + index;

            yield return g.First();
            index++;
        }
    }
}