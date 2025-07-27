namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelReader : IDisposable
{
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(bool useHeaderRow, string? sheetName, string startCell, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryAsync<T>(string? sheetName, string startCell, bool mapHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
    
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, string startCell, string endCell, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, string startCell, string endCell, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
    
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool useHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
}