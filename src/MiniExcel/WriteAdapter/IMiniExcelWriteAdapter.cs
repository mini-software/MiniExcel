using System.Collections.Generic;
using System.Threading;
using MiniExcelLibs.Utils;

namespace MiniExcelLibs.WriteAdapter;

internal interface IMiniExcelWriteAdapter
{
    bool TryGetKnownCount(out int count);
    List<ExcelColumnInfo>? GetColumns();
    IEnumerable<IEnumerable<CellWriteInfo>> GetRows(List<ExcelColumnInfo> props, CancellationToken cancellationToken = default);
}

internal readonly struct CellWriteInfo(object value, int cellIndex, ExcelColumnInfo prop)
{
    public object Value { get; } = value;
    public int CellIndex { get; } = cellIndex;
    public ExcelColumnInfo Prop { get; } = prop;
}