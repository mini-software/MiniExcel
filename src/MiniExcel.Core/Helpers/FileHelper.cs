namespace MiniExcelLib.Core.Helpers;

public static class FileHelper
{
    public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
}