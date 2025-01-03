using System;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace MiniExcelLibs
{
#if !NET8_0_OR_GREATER
    public interface IMiniExcelDataReader : IDataReader
#else
    public interface IMiniExcelDataReader : IDataReader, IAsyncDisposable
#endif
    {
        Task CloseAsync();

        Task<string> GetNameAsync(int i, CancellationToken cancellationToken = default);

        Task<object> GetValueAsync(int i, CancellationToken cancellationToken = default);

        Task<bool> NextResultAsync(CancellationToken cancellationToken = default);

        Task<bool> ReadAsync(CancellationToken cancellationToken = default);
    }
}
