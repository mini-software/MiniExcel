using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
    public interface IMiniExcelDataReader : IDataReader
    {
        Task CloseAsync();

        Task<string> GetNameAsync(int i);

        Task<string> GetNameAsync(int i, CancellationToken cancellationToken);

        Task<object> GetValueAsync(int i);

        Task<object> GetValueAsync(int i, CancellationToken cancellationToken);

        Task<bool> NextResultAsync();

        Task<bool> NextResultAsync(CancellationToken cancellationToken);

        Task<bool> ReadAsync();

        Task<bool> ReadAsync(CancellationToken cancellationToken);

        Task DisposeAsync();
    }
}
