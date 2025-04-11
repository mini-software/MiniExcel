using System;
using System.IO;

namespace MiniExcelLibs.Utils
{
    public static class ExcelTypeHelper
    {
        internal static ExcelType GetExcelType(string filePath, ExcelType excelType)
        {
            if (excelType != ExcelType.UNKNOWN)
                return excelType;
            
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            switch (extension)
            {
                case ".csv":
                    return ExcelType.CSV;
                case ".xlsx":
                case ".xlsm":
                    return ExcelType.XLSX;
                //case ".xls":
                //    return ExcelType.XLS;
                default:
                    throw new NotSupportedException($"Extension {extension} is not suppprted. Try specifying the ExcelType if you know what the underlying format is.");
            }
        }

        internal static ExcelType GetExcelType(Stream stream, ExcelType excelType)
        {
            if (excelType != ExcelType.UNKNOWN)
                return excelType;

            var probe = new byte[8];
            stream.Seek(0, SeekOrigin.Begin);
            stream.Read(probe, 0, probe.Length);
            stream.Seek(0, SeekOrigin.Begin);

            // New office format (can be any ZIP archive)
            if (probe[0] == 0x50 && probe[1] == 0x4B)
            {
                return ExcelType.XLSX;
            }

            throw new NotSupportedException("Stream cannot know the file type, please specify ExcelType manually");
        }
    }
}
