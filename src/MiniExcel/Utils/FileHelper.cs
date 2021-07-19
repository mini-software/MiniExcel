namespace MiniExcelLibs.Utils
{
    using System.IO;

    internal static partial class FileHelper
    {
        public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }

}
