using System.Collections.Generic;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName,string startCell, IConfiguration configuration);
        IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new();
    }

    internal interface IExcelReaderAsync : IExcelReader
    {
        Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, IConfiguration configuration);
        Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new();
    }
}
