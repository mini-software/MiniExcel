using System.IO;

namespace MiniExcelLibs
{
    internal interface IExcelWriter
    {
        void SaveAs(Stream stream, object value);
    }
}
