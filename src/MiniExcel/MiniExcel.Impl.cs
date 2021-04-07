namespace MiniExcelLibs
{
    using System;
    using System.IO;
    public static partial class MiniExcel
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
                    return ExcelType.XLSX;
                //case ".xls":
                //    return ExcelType.XLS;
                default:
                    throw new NotSupportedException($"Extension : {extension} not suppprt, or you can specify exceltype.");
            }
        }

        internal static ExcelType GetExcelType(Stream stream, ExcelType excelType)
        {
            if (excelType != ExcelType.UNKNOWN)
                return excelType;
            var buffer = new byte[512];
            stream.Read(buffer, 0, buffer.Length);
            var flag = BitConverter.ToUInt32(buffer, 0);
            stream.Position = 0;
            switch (flag)
            {
                // Old office format (can be any office file)
                //case 0xE011CFD0:
                //    return ExcelType.XLS;
                // New office format (can be any ZIP archive)
                case 0x04034B50:
                    return ExcelType.XLSX;
                default:
                    return ExcelType.CSV; //TODO:need to optimize
            }
        }
    }
}
