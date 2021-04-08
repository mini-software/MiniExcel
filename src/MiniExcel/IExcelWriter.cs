using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelWriter
    {
        void SaveAs(object value, bool printHeader, IConfiguration configuration);
    }
}
