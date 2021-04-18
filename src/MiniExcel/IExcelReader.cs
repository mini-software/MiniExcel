using System.Collections.Generic;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName, IConfiguration configuration);
        IEnumerable<T> Query<T>(string sheetName, IConfiguration configuration) where T : class, new();
    }
}
