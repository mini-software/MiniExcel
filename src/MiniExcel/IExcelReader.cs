using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(Stream stream, bool UseHeaderRow = false);
        IEnumerable<T> Query<T>(Stream stream) where T : class, new();
    }

    internal interface IExcelWriter {
        void SaveAs(Stream stream, object value);
    }
}
