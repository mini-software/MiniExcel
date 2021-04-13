namespace MiniExcelLibs
{
    using MiniExcelLibs.OpenXml;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.IO.Compression;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;

    public static partial class MiniExcel
    {
        public static void SaveAs(string path, object value, bool printHeader = true, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                SaveAs(stream, value, printHeader, sheetName, GetExcelType(path, excelType), configuration);
        }

        /// <summary>
        /// Default SaveAs Xlsx file
        /// </summary>
        public static void SaveAs(this Stream stream, object value, bool printHeader = true, string sheetName = null, ExcelType excelType = ExcelType.XLSX, IConfiguration configuration = null)
        {
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            ExcelWriterFactory.GetProvider(stream, excelType).SaveAs(value, printHeader, configuration);
        }

        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query<T>(stream, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item; //Foreach yield return twice reason : https://stackoverflow.com/questions/66791982/ienumerable-extract-code-lazy-loading-show-stream-was-not-readable
        }

        public static IEnumerable<T> Query<T>(this Stream stream, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null) where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query<T>(sheetName, configuration);
        }

        public static IEnumerable<dynamic> Query(string path, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in Query(stream, useHeaderRow, sheetName, GetExcelType(path, excelType), configuration))
                    yield return item;
        }

        public static IEnumerable<dynamic> Query(this Stream stream, bool useHeaderRow = false, string sheetName = null, ExcelType excelType = ExcelType.UNKNOWN, IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream, GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, configuration);
        }

        public static IEnumerable<string> GetSheetNames(string path)
        {
            using (var stream = File.OpenRead(path))
                foreach (var item in GetSheetNames(stream))
                    yield return item;
        }

        public static IEnumerable<string> GetSheetNames(this Stream stream)
        {
            var archive = new ExcelOpenXmlZip(stream);
            foreach (var item in ExcelOpenXmlSheetReader.GetWorkbookRels(archive.Entries))
                yield return item.Name;
        }

        public static void SaveAsByTemplate(string path, string templatePath, object value)
        {
            using(var stream = File.Create(path))
                SaveAsByTemplateImpl(stream, templatePath, value);
        }

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
        {
            SaveAsByTemplateImpl(stream, templatePath, value);
        }

        internal static void SaveAsByTemplateImpl(Stream stream, string templatePath, object value)
        {
            //only support xlsx         
            var values = (Dictionary<string, object>)value;

            //TODO: copy new bytes 
            using (var templateStream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                templateStream.CopyTo(stream);

                var reader = new ExcelOpenXmlSheetReader(stream);
                var _archive = new ExcelOpenXmlZip(stream, mode: ZipArchiveMode.Update, false, Encoding.UTF8);
                {
                    //TODO: read sharedString
                    var sharedStrings = reader.GetSharedStrings();

                    //TODO: read all xlsx sheets
                    var sheets = _archive.ZipFile.Entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                         || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                    foreach (var sheet in sheets)
                    {
                        var sheetStream = sheet.Open();

                        var doc = new System.Xml.XmlDocument();
                        doc.Load(sheetStream);

                        XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
                        ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                        var worksheet = doc.SelectSingleNode("/x:worksheet", ns);
                        var isPrefix = worksheet.Name.Contains(":");

                        var rows = doc.SelectNodes($"/x:worksheet/x:sheetData/x:row", ns);

                        foreach (XmlElement row in rows)
                        {
                            var cs = row.SelectNodes($"x:c", ns);
                            foreach (XmlElement c in cs)
                            {
                                var t = c.GetAttribute("t");
                                var v = c.SelectSingleNode("x:v", ns);
                                if (v == null || v.InnerText == null) //![image](https://user-images.githubusercontent.com/12729184/114363496-075a3f80-9bab-11eb-9883-8e3fec10765c.png)
                                    continue;

                                if (t == "s")
                                {
                                    //need to check sharedstring not exist
                                    if (sharedStrings.ElementAtOrDefault(int.Parse(v.InnerText)) != null)
                                    {
                                        v.InnerText = sharedStrings[int.Parse(v.InnerText)];
                                        // change type = str and replace its value
                                        c.SetAttribute("t", "str");
                                    }
                                    //TODO: remove sharedstring 
                                }
                            }
                        }
                        sheetStream.Dispose();

                        var fullName = sheet.FullName;
                        sheet.Delete();
                        ZipArchiveEntry entry = _archive.ZipFile.CreateEntry(fullName);
                        using (var zipStream = entry.Open())
                        {
                            ExcelOpenXmlTemplate.GenerateSheetXml(zipStream, doc.InnerXml, values);
                            //doc.Save(zipStream); //don't do it beacause : ![image](https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png)
                        }
                    }
                }

                _archive.Dispose();
            }
        }

        private static Type GetIEnumerableRuntimeValueType(object v)
        {
            if (v == null)
                throw new InvalidDataException("input parameter value can't be null");
            foreach (var tv in v as IEnumerable)
            {
                if (tv != null)
                {
                    return tv.GetType();
                }
            }
            throw new InvalidDataException("can't get parameter type information");
        }
    }
}
