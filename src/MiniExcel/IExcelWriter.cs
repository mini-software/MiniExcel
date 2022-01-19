using System.IO;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelWriter 
    {
        void SaveAs();
        Task SaveAsAsync();
    }
}
