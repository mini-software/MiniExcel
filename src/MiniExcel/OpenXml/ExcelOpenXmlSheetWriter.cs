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
        internal static void SaveAs(string path, object value, bool printHeader)
        {
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

        internal static void SaveAs(Stream stream, object value, bool printHeader)
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

        internal static void GenerateSheet(StreamWriter writer, ZipArchive archive,object value, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            // dimension
            var dimensionRef = "";
            var maxRowIndex = 0;
            var maxColumnIndex = 0;


            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                if (value is DataTable)
                {
                    var dt = value as DataTable;

                    // dimension
                    maxRowIndex = dt.Rows.Count + (printHeader && dt.Rows.Count >0 ? 1 :0);
                    maxColumnIndex = dt.Columns.Count;
                    writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");


                    if (printHeader)
                    {
                        writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;
                        foreach (DataColumn c in dt.Columns)
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

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                        var xIndex = xy.Item1;

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            var cellValue = dt.Rows[i][j];
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
                else if (IsDapperRowOrDictionaryStringObject(value))
                {
                    var collection = value as IEnumerable<IDictionary<string, object>>;




                    var firstTime = true;
                    ICollection<string> keys = null;
                    foreach (var v in collection)
                    {
                        // head
                        if (firstTime)
                        {
                            firstTime = false;
                            if (v == null)
                                continue;

                            keys = v.Keys;

                            
                            // dimension 
                            maxColumnIndex = keys.Count;
                            var cnt = collection.Count();
                            maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
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
                else if (value is ICollection)
                {
                    var collection = value as ICollection;

                    var props = Helpers.GetSubtypeProperties(collection);

                    maxColumnIndex = props.Length;
                    var cnt = collection.Count;
                    maxRowIndex = cnt + (printHeader && cnt > 0 ? 1 : 0);  //TODO:it can optimize
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

                    foreach (var v in collection)
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
                else
                {
                    throw new NotImplementedException($"{value?.GetType().Name} type not implemented,please issue for me : https://github.com/shps951023/MiniExcel/issues");
                }
                writer.Write("</x:sheetData></x:worksheet>");
            }
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
