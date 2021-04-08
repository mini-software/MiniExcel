using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName);
        IEnumerable<T> Query<T>(string sheetName) where T : class, new();
    }
}
