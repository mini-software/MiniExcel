namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelWriteAdapterAsync 
{
    Task<List<MiniExcelColumnMapping>?> GetColumnsAsync();
    IAsyncEnumerable<CellWriteInfo[]> GetRowsAsync(List<MiniExcelColumnMapping> mappings, CancellationToken cancellationToken);
}