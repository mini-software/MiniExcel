﻿using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelReader : IDisposable
    {
        IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell);
        IEnumerable<T> Query<T>(string sheetName, string startCell, bool hasHeader) where T : class, new();
        Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default);
        Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new();
        IEnumerable<IDictionary<string, object>> QueryRange(bool useHeaderRow, string sheetName, string startCell, string endCell);
        IEnumerable<T> QueryRange<T>(string sheetName, string startCell, string endCell) where T : class, new();
        Task<IEnumerable<IDictionary<string, object>>> QueryAsyncRange(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default);
        Task<IEnumerable<T>> QueryAsyncRange<T>(string sheetName, string startCell, string endCell, bool hasHeader, CancellationToken cancellationToken = default) where T : class, new();
    }
}
