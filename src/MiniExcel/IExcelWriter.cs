using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelWriter 
    {
        void SaveAs();
        Task SaveAsAsync(CancellationToken cancellationToken = default(CancellationToken));
        void Insert();
    }
}
