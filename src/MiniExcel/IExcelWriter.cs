using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    internal interface IExcelWriter
    {
        void SaveAs();
        Task SaveAsAsync(CancellationToken cancellationToken = default);
        void Insert(bool overwriteSheet = false);
        Task InsertAsync(bool overwriteSheet = false, CancellationToken cancellationToken = default);
    }
}
