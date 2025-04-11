using System.IO;

namespace MiniExcelLibs.Utils
{
    internal static class FileHelper
    {
        public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }
}
