/**
 This Class Modified from ExcelDataReader : https://github.com/ExcelDataReader/ExcelDataReader
 **/
namespace MiniExcelLibs.Tests.Utils
{
    using System;
    using System.IO;

    internal static class PathHelper
    {
        public static string GetSamplePath(string fileName)
        {
            return $@"../../../../../samples/{fileName}";
        }

        public static string GetNewTemplateFilePath(string extension=".xlsx") => 
            Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.{extension}");

    }

}
