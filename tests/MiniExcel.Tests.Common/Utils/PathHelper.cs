namespace MiniExcelLib.Tests.Common.Utils;

public static class PathHelper
{
    public static string GetFile(string fileName) => $"../../../../../samples/{fileName}";

    public static string GetTempPath(string extension = "xlsx")
    {
        var method = new System.Diagnostics.StackTrace().GetFrame(1)?.GetMethod();

        var path = Path.Combine(Path.GetTempPath(), $"{method?.DeclaringType?.Name}_{method?.Name}.{extension}")
            .Replace("<", string.Empty)
            .Replace(">", string.Empty);
        
        if (File.Exists(path))
            File.Delete(path);
        
        return path;
    }

    public static string GetTempFilePath(string extension = "xlsx") => $"{Path.GetTempPath()}{Guid.NewGuid()}.{extension}";
}