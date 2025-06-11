using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlSheetReader : IExcelReader
    {
        public async Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool useHeaderRow, string sheetName, string startCell, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => Query(useHeaderRow, sheetName, startCell), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => Query<T>(sheetName, startCell, hasHeader), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, string startCell, string endCell, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => Query(useHeaderRow, sheetName, startCell), cancellationToken).ConfigureAwait(false);
        }

        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, string startCell, string endCell, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startCell, endCell, hasHeader), cancellationToken).ConfigureAwait(false);
        }
               
        public async Task<IEnumerable<IDictionary<string, object>>> QueryRangeAsync(bool useHeaderRow, string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => QueryRange(useHeaderRow, sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex), cancellationToken).ConfigureAwait(false);
        }
        
        public async Task<IEnumerable<T>> QueryRangeAsync<T>(string sheetName, int startRowIndex, int startColumnIndex, int? endRowIndex, int? endColumnIndex, bool hasHeader = true, CancellationToken cancellationToken = default) where T : class, new()
        {
            return await Task.Run(() => QueryRange<T>(sheetName, startRowIndex, startColumnIndex, endRowIndex, endColumnIndex, hasHeader), cancellationToken).ConfigureAwait(false);
        }
    }
}
