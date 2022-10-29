using System;
using MiniExcelLibs.Utils;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelReader: IDisposable
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName,string startCell);
        IEnumerable<T> Query<T>(string sheetName, string startCell) where T : class, new();
        Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell,CancellationToken cancellationToken = default(CancellationToken));
        Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell,CancellationToken cancellationToken = default(CancellationToken)) where T : class, new();
        IEnumerable<IDictionary<string, object>> QueryRange(bool UseHeaderRow, string sheetName, string startCell, string endCell);
        IEnumerable<T> QueryRange<T>(string sheetName, string startCell, string endCell) where T : class, new();
        Task<IEnumerable<IDictionary<string, object>>> QueryAsyncRange(bool UseHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default(CancellationToken));
        Task<IEnumerable<T>> QueryAsyncRange<T>(string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default(CancellationToken)) where T : class, new();
    }
}
