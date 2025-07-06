namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelDataReader : IDataReader
#if NET8_0_OR_GREATER
    ,IAsyncDisposable
#endif
{
    Task CloseAsync();
    Task<string> GetNameAsync(int i, CancellationToken cancellationToken = default);
    Task<object?> GetValueAsync(int i, CancellationToken cancellationToken = default);
    Task<bool> NextResultAsync(CancellationToken cancellationToken = default);
    Task<bool> ReadAsync(CancellationToken cancellationToken = default);
}