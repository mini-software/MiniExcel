namespace MiniExcelLibs.Tests.Utils;

public class AutoDeletingPath : IDisposable
{
    public string FilePath { get; }

    private AutoDeletingPath(string path)
    {
        FilePath = path;
    }

    public static AutoDeletingPath Create(string path) => new(path);
    public static AutoDeletingPath Create(string path, string filename) => new(Path.Combine(path, filename));
    public static AutoDeletingPath Create(ExcelType type = ExcelType.XLSX) => Create(
        Path.GetTempPath(), 
        $"{Guid.NewGuid()}.{type.ToString().ToLowerInvariant()}");

    public void Dispose()
    {
        File.Delete(FilePath);
    }
    
    public override string ToString() => FilePath;
}