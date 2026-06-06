namespace MiniExcelLib.Core.Abstractions;

public partial interface IMiniExcelReader : IDisposable, IAsyncDisposable
{
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryAsync(bool hasHeaderRow, string? sheetName, string startCell, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryAsync<T>(string? sheetName, string startCell, bool mapHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
    
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool hasHeaderRow, string? sheetName, string startCell, string endCell, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, string startCell, string endCell, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
    
    [CreateSyncVersion]
    IAsyncEnumerable<IDictionary<string, object?>> QueryRangeAsync(bool hasHeaderRow, string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default);
    
    [CreateSyncVersion]
    IAsyncEnumerable<T> QueryRangeAsync<T>(string? sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool treatHeaderAsData, CancellationToken cancellationToken = default) where T : class, new();
}