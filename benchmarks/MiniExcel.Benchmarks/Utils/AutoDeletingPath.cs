namespace MiniExcelLibs.Benchmarks.Utils;

internal class AutoDeletingPath : IDisposable
{
    public string FilePath { get; private set; }

    private AutoDeletingPath(string path)
    {
        FilePath = path;
    }

    internal static AutoDeletingPath Create(string path) => new(path);
    internal static AutoDeletingPath Create(string path, string filename) => new(Path.Combine(path, filename));
    internal static AutoDeletingPath Create() => Create(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");

    public void Dispose()
    {
        File.Delete(FilePath);
    }
}
