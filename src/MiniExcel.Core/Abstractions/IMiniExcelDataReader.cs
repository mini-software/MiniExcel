namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelDataReader : IDataReader
#if NET8_0_OR_GREATER
    ,IAsyncDisposable
#endif
{
    Task CloseAsync();
    Task<bool> ReadAsync(CancellationToken cancellationToken = default);
}