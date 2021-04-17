using System.Collections.Generic;
using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName, IConfiguration configuration);
        IEnumerable<T> Query<T>(string sheetName, IConfiguration configuration) where T : class, new();
    }

    internal interface IExcelTemplate
    {
        //TODO: add byte or stream templatePath
        void SaveAsByTemplate(string templatePath, object value);
        void SaveAsByTemplate(byte[] templateBtyes, object value);
    }
}
