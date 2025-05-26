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
            var read = stream.Read(probe, 0, probe.Length);
            if (read != probe.Length)
                throw new InvalidDataException("The file/stream does not contain enough data to process");
            
            stream.Seek(0, SeekOrigin.Begin);

            // New office format (can be any ZIP archive)
            if (probe[0] == 0x50 && probe[1] == 0x4B)
            {
                return ExcelType.XLSX;
            }

            throw new InvalidDataException("The file type could not be inferred automatically, please specify ExcelType manually");
        }
    }
}
