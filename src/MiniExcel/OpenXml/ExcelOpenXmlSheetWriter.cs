using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlSheetWriter
    {
        internal static void SaveAs(string path, DataTable value, bool printHeader)
        {
            //StreamWriter writer, ZipArchive archive,object value, bool printHeader
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var defaults = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    defaults.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, defaults);
            }
        }

        internal static void SaveAs<T>(string path, ICollection<T> value, bool printHeader)
        {
            //StreamWriter writer, ZipArchive archive,object value, bool printHeader
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var defaults = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    defaults.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, defaults);
            }
        }

        internal static void SaveAs<T>(string path, IEnumerable<T> value, bool printHeader)
        {
            //StreamWriter writer, ZipArchive archive,object value, bool printHeader
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var defaults = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    defaults.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, defaults);
            }
        }

        internal static void SaveAs(Stream stream, DataTable value, bool printHeader)
        {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var packages = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    packages.Add(sheetPath,new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, packages);
            }
        }

        internal static void SaveAs<T>(Stream stream, ICollection<T> value, bool printHeader)
        {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var packages = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    packages.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, packages);
            }
        }

        internal static void SaveAs<T>(Stream stream, IEnumerable<T> value, bool printHeader)
        {
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, true, Utf8WithBom))
            {
                var packages = DefualtOpenXml.GenerateDefaultOpenXml(archive);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                    {
                        GenerateSheet(writer, archive, value, printHeader);
                    }
                    packages.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, packages);
            }
        }

        internal static void GenerateSheet<T>(StreamWriter writer, ZipArchive archive, ICollection<T> value, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");
            var yIndex = xy.Item2;

            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");

            var cnt = value.Count;
            if (value == null || cnt == 0)
            {
                if (printHeader)
                {
                    var props = Helpers.GetProperties(typeof(T));
                    var maxColumnIndex = props.Length;
                    var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                    writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                        writer.Write($"<x:v>{p.Name}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    writer.Write("</x:sheetData></x:worksheet>");
                }
                else
                    writer.Write($@"<dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");

                return;
            }

            if (Helpers.IsAssignableFromIDictionary<T>())
            {
                var firstTime = true;
                ICollection keys = null;
                foreach (IDictionary v in value)
                {
                    // head
                    if (firstTime)
                    {
                        firstTime = false;
                        if (v == null)
                            continue;
                        keys = v.Keys;


                        // dimension 
                        var maxColumnIndex = keys.Count;
                        var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                        writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");


                        if (printHeader)
                        {
                            writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                            var xIndex = xy.Item1;
                            foreach (var key in keys)
                            {
                                var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                                writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                                writer.Write($"<x:v>{key}");
                                writer.Write($"</x:v>");
                                writer.Write($"</x:c>");
                                xIndex++;
                            }
                            writer.Write($"</x:row>");
                            yIndex++;
                        }
                    }

                    //body
                    {
                        writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in keys)
                        {
                            var cellValue = v[p];
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            var t = "t=\"str\"";
                            {
                                if (decimal.TryParse(cellValueStr, out var outV))
                                    t = "t=\"n\"";
                                if (cellValue is bool)
                                {
                                    t = "t=\"b\"";
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                if (cellValue is DateTime || cellValue is DateTime?)
                                {
                                    t = "s=\"1\"";
                                    cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                                }
                            }
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            writer.Write($"<x:c r=\"{columname}\" {t}>");
                            writer.Write($"<x:v>{cellValueStr}");
                            writer.Write($"</x:v>");
                            writer.Write($"</x:c>");
                            xIndex++;
                        }
                        writer.Write($"</x:row>");
                        yIndex++;
                    }
                }
            }
            else
            {
                // demension
                var props = Helpers.GetProperties(typeof(T));
                var maxColumnIndex = props.Length;
                var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

                if (printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                        writer.Write($"<x:v>{p.Name}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }

                foreach (var v in value)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var cellValue = p.GetValue(v);
                        var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                        var t = "t=\"str\"";
                        {
                            if (decimal.TryParse(cellValueStr, out var outV))
                                t = "t=\"n\"";
                            if (cellValue is bool)
                            {
                                t = "t=\"b\"";
                                cellValueStr = (bool)cellValue ? "1" : "0";
                            }
                            if (cellValue is DateTime || cellValue is DateTime?)
                            {
                                t = "s=\"1\"";
                                cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                            }
                        }
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" {t}>");
                        writer.Write($"<x:v>{cellValueStr}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
            }
            writer.Write("</x:sheetData></x:worksheet>");
        }

        internal static void GenerateSheet<T>(StreamWriter writer, ZipArchive archive, IEnumerable<T> value, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            var yIndex = xy.Item2;

            var cnt = value.Count();
            if (value == null || cnt == 0)
            {
                if (printHeader)
                {
                    var props = Helpers.GetProperties(typeof(T));
                    var maxColumnIndex = props.Length;
                    var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                    writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                        writer.Write($"<x:v>{p.Name}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    writer.Write("</x:sheetData></x:worksheet>");
                }
                else
                {
                    writer.Write($@"<dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");
                }
                return;
            }

            if (value is IEnumerable<IDictionary<string,object>>)
            {
                var firstTime = true;
                ICollection<string> keys = null;
                foreach (IDictionary<string,object> v in value)
                {
                    // head
                    if (firstTime)
                    {
                        firstTime = false;
                        if (v == null)
                            continue;
                        keys = v.Keys;

                        // dimension 
                        var maxColumnIndex = keys.Count;

                        var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                        writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");


                        if (printHeader)
                        {
                            writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                            var xIndex = xy.Item1;
                            foreach (var key in keys)
                            {
                                var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                                writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                                writer.Write($"<x:v>{key}");
                                writer.Write($"</x:v>");
                                writer.Write($"</x:c>");
                                xIndex++;
                            }
                            writer.Write($"</x:row>");
                            yIndex++;
                        }
                    }

                    //body
                    {
                        writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in keys)
                        {
                            var cellValue = v[p];
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            var t = "t=\"str\"";
                            {
                                if (decimal.TryParse(cellValueStr, out var outV))
                                    t = "t=\"n\"";
                                if (cellValue is bool)
                                {
                                    t = "t=\"b\"";
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                if (cellValue is DateTime || cellValue is DateTime?)
                                {
                                    t = "s=\"1\"";
                                    cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                                }
                            }
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            writer.Write($"<x:c r=\"{columname}\" {t}>");
                            writer.Write($"<x:v>{cellValueStr}");
                            writer.Write($"</x:v>");
                            writer.Write($"</x:c>");
                            xIndex++;
                        }
                        writer.Write($"</x:row>");
                        yIndex++;
                    }
                }
            }
            else if (Helpers.IsAssignableFromIDictionary<T>())
            {
                var firstTime = true;
                ICollection keys = null;
                foreach (IDictionary v in value)
                {
                    // head
                    if (firstTime)
                    {
                        firstTime = false;
                        if (v == null)
                            continue;
                        keys = v.Keys;

                        // dimension 
                        var maxColumnIndex = keys.Count;
                        
                        var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                        writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");


                        if (printHeader)
                        {
                            writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                            var xIndex = xy.Item1;
                            foreach (var key in keys)
                            {
                                var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                                writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                                writer.Write($"<x:v>{key}");
                                writer.Write($"</x:v>");
                                writer.Write($"</x:c>");
                                xIndex++;
                            }
                            writer.Write($"</x:row>");
                            yIndex++;
                        }
                    }

                    //body
                    {
                        writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (var p in keys)
                        {
                            var cellValue = v[p];
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            var t = "t=\"str\"";
                            {
                                if (decimal.TryParse(cellValueStr, out var outV))
                                    t = "t=\"n\"";
                                if (cellValue is bool)
                                {
                                    t = "t=\"b\"";
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                if (cellValue is DateTime || cellValue is DateTime?)
                                {
                                    t = "s=\"1\"";
                                    cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                                }
                            }
                            var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                            writer.Write($"<x:c r=\"{columname}\" {t}>");
                            writer.Write($"<x:v>{cellValueStr}");
                            writer.Write($"</x:v>");
                            writer.Write($"</x:c>");
                            xIndex++;
                        }
                        writer.Write($"</x:row>");
                        yIndex++;
                    }
                }
            }
            else
            {
                // demension
                var props = Helpers.GetProperties(typeof(T));
                var maxColumnIndex = props.Length;
                var maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
                writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

                if (printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                        writer.Write($"<x:v>{p.Name}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }

                foreach (var v in value)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (var p in props)
                    {
                        var cellValue = p.GetValue(v);
                        var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                        var t = "t=\"str\"";
                        {
                            if (decimal.TryParse(cellValueStr, out var outV))
                                t = "t=\"n\"";
                            if (cellValue is bool)
                            {
                                t = "t=\"b\"";
                                cellValueStr = (bool)cellValue ? "1" : "0";
                            }
                            if (cellValue is DateTime || cellValue is DateTime?)
                            {
                                t = "s=\"1\"";
                                cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                            }
                        }
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" {t}>");
                        writer.Write($"<x:v>{cellValueStr}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
            }

            writer.Write("</x:sheetData></x:worksheet>");
        }

        internal static void GenerateSheet(StreamWriter writer, ZipArchive archive,DataTable value, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                // dimension
                var maxRowIndex = value.Rows.Count + (printHeader && value.Rows.Count >0 ? 1 :0);
                var maxColumnIndex = value.Columns.Count;
                writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

                if (printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (DataColumn c in value.Columns)
                    {
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" t=\"str\">");
                        writer.Write($"<x:v>{c.ColumnName}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }

                for (int i = 0; i < value.Rows.Count; i++)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;

                    for (int j = 0; j < value.Columns.Count; j++)
                    {
                        var cellValue = value.Rows[i][j];
                        var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                        var t = "t=\"str\"";
                        {
                            if (decimal.TryParse(cellValueStr, out var outV))
                                t = "t=\"n\"";
                            if (cellValue is bool)
                            {
                                t = "t=\"b\"";
                                cellValueStr = (bool)cellValue ? "1" : "0";
                            }
                            if (cellValue is DateTime || cellValue is DateTime?)
                            {
                                t = "s=\"1\"";
                                cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
                            }
                        }
                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{columname}\" {t}>");
                        writer.Write($"<x:v>{cellValueStr}");
                        writer.Write($"</x:v>");
                        writer.Write($"</x:c>");
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
            }
            writer.Write("</x:sheetData></x:worksheet>");
        }

        private static void GenerateContentTypesXml(ZipArchive archive, Dictionary<string, ZipPackageInfo> defaults)
        {
            //[Content_Types].xml 

            var sb = new StringBuilder(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default ContentType=""application/xml"" Extension=""xml""/><Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>");
            foreach (var p in defaults)
                sb.Append($"<Override ContentType=\"{p.Value.ContentType}\" PartName=\"/{p.Key}\" />");
            sb.Append("</Types>");

            ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml");
            using (var zipStream = entry.Open())
            using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                writer.Write(sb.ToString());
        }

        private static string GetDimension(int maxRowIndex, int maxColumnIndex)
        {
            string dimensionRef;
            if (maxRowIndex == 0 && maxColumnIndex == 0)
                dimensionRef = "A1";
            else if (maxColumnIndex == 1)
                dimensionRef = $"A{maxRowIndex}";
            else if (maxRowIndex == 0)
                dimensionRef = $"A1:{Helpers.GetAlphabetColumnName(maxColumnIndex - 1)}1";
            else
                dimensionRef = $"A1:{Helpers.GetAlphabetColumnName(maxColumnIndex - 1)}{maxRowIndex}";
            return dimensionRef;
        }

        private readonly static UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);

        private static bool IsDapperRowOrDictionaryStringObject(object value)
        {
            return value is IEnumerable<IDictionary<string, object>>;
        }
    }
}
