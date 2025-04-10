﻿namespace MiniExcelLibs.Tests.Utils;

public static class FileHelper
{
    public static Stream OpenRead(string path)
    {
        try
        {
            return File.OpenRead(path);
        }
        catch (IOException)
        {
            var newPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
            File.Copy(path, newPath);
            return File.OpenRead(newPath);
        }
    }
}