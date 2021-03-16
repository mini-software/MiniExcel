using System;
using System.IO;

namespace MiniExcelLibs.Tests
{
    public partial class MiniExcelHelperTests
    {
        public class FileHelper
        {
            public static Stream OpenRead(string path)
            {
                try
                {
                    return File.OpenRead(path);
                }
                catch (System.IO.IOException)
                {
                    var newPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
                    File.Copy(path, newPath);
                    return File.OpenRead(newPath);
                }
            }

        }
    }
}