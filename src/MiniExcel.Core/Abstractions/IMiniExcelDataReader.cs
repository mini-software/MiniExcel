namespace MiniExcelLib.Core.Abstractions;

public interface IMiniExcelDataReader : IDataReader, IAsyncDisposable
{
    Task CloseAsync();
    Task<bool> ReadAsync(CancellationToken cancellationToken = default);
}