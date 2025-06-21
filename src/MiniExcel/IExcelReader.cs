using System;
using System.Collections.Generic;
using System.Threading;

namespace MiniExcelLibs
{
    internal partial interface IExcelReader : IDisposable
    {
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<IDictionary<string, object>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default);
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<T> QueryAsync<T>(string sheetName, string startCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new();
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<IDictionary<string, object>> QueryRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default);
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<T> QueryRangeAsync<T>(string sheetName, string startCell, string endCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new();
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<IDictionary<string, object>> QueryRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default);
        [Zomp.SyncMethodGenerator.CreateSyncVersion]
        IAsyncEnumerable<T> QueryRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new();
    }
}
