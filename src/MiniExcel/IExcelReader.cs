using System.Collections.Generic;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName,string startCell, IConfiguration configuration);
        IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new();
    }
}
