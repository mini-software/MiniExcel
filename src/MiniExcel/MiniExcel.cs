namespace MiniExcelLibs
{
    using System.Collections.Generic;
    using System.IO;

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, ExcelType excelType = ExcelType.UNKNOWN)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, GetExcelType(path, excelType));
        }

        /// <summary>
        /// Default SaveAs Xlsx file
        /// </summary>
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, ExcelType excelType = ExcelType.XLSX)
        {
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            ExcelFacorty.GetExcelProvider(excelType, printHeader).SaveAs(stream, value);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query<T>(stream, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            return ExcelFacorty.GetExcelProvider(GetExcelType(stream, excelType)).Query<T>(stream);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query(stream, useHeaderRow, GetExcelType(path, excelType), configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return ExcelFacorty.GetExcelProvider(GetExcelType(stream, excelType), useHeaderRow).Query(stream, useHeaderRow);
        }
    }
}
