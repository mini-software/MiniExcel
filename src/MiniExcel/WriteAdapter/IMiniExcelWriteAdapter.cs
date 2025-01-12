using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Threading;

namespace MiniExcelLibs.WriteAdapter
{
    internal interface IMiniExcelWriteAdapter
    {
        bool TryGetNonEnumeratedCount(out int count);

        List<ExcelColumnInfo> GetColumns();

        IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props, CancellationToken cancellationToken = default);
    }

    internal readonly struct CellWriteInfo
    {
        public CellWriteInfo(object value, int cellIndex, ExcelColumnInfo prop)
        {
            Value = value;
            CellIndex = cellIndex;
            Prop = prop;
        }

        public object Value { get; }
        public int CellIndex { get; }
        public ExcelColumnInfo Prop { get; }
    }
}


