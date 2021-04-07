using System.Collections.Generic;
using System.IO;
namespace MiniExcelLibs
{
    internal abstract class ExcelProviderBase : IExcelReader, IExcelWriter
    {
        public abstract IEnumerable<IDictionary<string, object>> Query(Stream stream, bool UseHeaderRow = false);
        public abstract IEnumerable<T> Query<T>(Stream stream) where T : class, new();
        public abstract void SaveAs(Stream stream, object input);
    }
}
