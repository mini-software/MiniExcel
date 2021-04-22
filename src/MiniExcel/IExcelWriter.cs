using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelWriter
    {
        void SaveAs(object value,string sheetName, bool printHeader, IConfiguration configuration);
    }
}
