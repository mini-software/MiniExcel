using MiniExcelLibs.Attributes;
using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Linq;

namespace MiniExcelLibs.OpenXml.Styles {
    public static class SheetStyleBuilderHelper
    {
        public static IEnumerable<ExcelColumnAttribute> GenerateStyleIds( int startUpCellXfs, ICollection<ExcelColumnAttribute> dynamicColumns ) {
            if ( dynamicColumns == null )
                yield break;

            int index = 0;
            foreach ( var g in dynamicColumns?.Where( x => !string.IsNullOrWhiteSpace( x.Format ) && new ExcelNumberFormat( x.Format ).IsValid ).GroupBy( x => x.Format ) ) {
                foreach ( var col in g )
                    col.FormatId = startUpCellXfs + index;

                yield return g.First();
                index++;
            }
        }
    }

}
