using System.IO;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelWriter 
    {
        void SaveAs(object value,string sheetName);
        Task SaveAsAsync(object value, string sheetName);
    }
}
