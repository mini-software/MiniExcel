namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using System.Linq;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.IO;
    using System.IO.Compression;
    using System.Text;
    using System.Reflection;
    using MiniExcelLibs.Utils;
    using System.Globalization;
    using System.Collections;

    public static partial class MiniExcel
    {
        private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);

        public static void SaveAs(this Stream stream, object value, string startCell = "A1", bool printHeader = true)
        {
            SaveAsImpl(stream, GetCreateXlsxInfos(value, startCell, printHeader));
            stream.Position = 0;
        }

        public static void SaveAs(string filePath, object value, string startCell = "A1", bool printHeader = true)
        {
            SaveAsImpl(filePath, GetCreateXlsxInfos(value, startCell, printHeader));
        }

        public static IEnumerable<T> Query<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream);
        }

        public static T QueryFirst<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).First();
        }

        public static T QueryFirstOrDefault<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).FirstOrDefault();
        }

        public static T QuerySingle<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).Single();
        }

        public static T QuerySingleOrDefault<T>(this Stream stream) where T : class, new()
        {
            return QueryImpl<T>(stream).SingleOrDefault();
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow);
        }

        public static dynamic QueryFirst(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).First();
        }

        public static dynamic QueryFirstOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).FirstOrDefault();
        }

        public static dynamic QuerySingle(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).Single();
        }

        public static dynamic QuerySingleOrDefault(this Stream stream, bool useHeaderRow = false)
        {
            return new ExcelOpenXmlSheetReader().QueryImpl(stream, useHeaderRow).SingleOrDefault();
        }
    }
}
