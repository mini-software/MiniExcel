using MiniExcelLib.Core.Attributes;

namespace MiniExcelLib.OpenXml.Styles.Builder;

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
        
        foreach (var group in cols) 
        {
            foreach (var col in group)
            {
                col.SetFormatId(startUpCellXfs + index);
            }

            yield return group.First();
            index++;
        }
    }
}