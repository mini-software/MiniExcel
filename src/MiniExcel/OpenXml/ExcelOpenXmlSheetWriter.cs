using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly bool printHeader;
        public ExcelOpenXmlSheetWriter(bool printHeader)
        {
            this.printHeader = printHeader;
        }

        public void SaveAs(Stream stream, object value)
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
                        if (value == null)
                        {
                            WriteEmptySheet(writer);
                            goto End;
                        }

                        var type = value.GetType();

                        Type genericType = null;

                        //DapperRow
                        if (value is IEnumerable)
                        {
                            var values = value as IEnumerable;

                            var rowCount = 0;

                            var maxColumnIndex = 0;
                            List<object> keys = new List<object>();
                            PropertyInfo[] props = null;
                            string mode = null;

                            {
                                foreach (var item in values) //TODO: need to optimize
                                {
                                    rowCount = checked(rowCount + 1);
                                    if (item != null && mode == null)
                                    {
                                        if (item is IDictionary<string, object>)
                                        {
                                            var item2 = item as IDictionary<string, object>;
                                            mode = "IDictionary<string, object>";
                                            maxColumnIndex = item2.Keys.Count;
                                            foreach (var key in item2.Keys)
                                                keys.Add(key);
                                        }
                                        else if (item is IDictionary)
                                        {
                                            var item2 = item as IDictionary;
                                            mode = "IDictionary";
                                            maxColumnIndex = item2.Keys.Count;
                                            foreach (var key in item2.Keys)
                                                keys.Add(key);
                                        }
                                        else
                                        {
                                            mode = "Properties";
                                            genericType = item.GetType();
                                            props = Helpers.GetProperties(genericType);
                                            //props = genericType.GetProperties();
                                            if (props.Length == 0)
                                                throw new InvalidOperationException($"Generic Type : {genericType} valid properties count is 0, if you have trouble please issue for me.");
                                            maxColumnIndex = props.Length;
                                        }

                                        // not re-foreach key point
                                        var collection = value as ICollection;
                                        if (collection != null)
                                        {
                                            rowCount = checked((value as ICollection).Count);
                                            break;
                                        }
                                        continue;
                                    }
                                }
                            }

                            if (rowCount == 0)
                            {
                                WriteEmptySheet(writer);
                                goto End;
                            }

                            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
                            // dimension 

                            var maxRowIndex = rowCount + (printHeader && rowCount > 0 ? 1 : 0);  //TODO:it can optimize
                            writer.Write($@"<dimension ref=""{GetDimension(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

                            //header
                            var yIndex = 1;
                            var xIndex = 1;
                            if (printHeader)
                            {
                                var cellIndex = xIndex;
                                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                                if (props != null)
                                {
                                    foreach (var p in props)
                                    {
                                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                                        writer.Write($"<x:c r=\"{columname}\" t=\"str\"><x:v>{p.Name}</x:v></x:c>");
                                        cellIndex++;
                                    }
                                }
                                else
                                {
                                    foreach (var key in keys)
                                    {
                                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                                        writer.Write($"<x:c r=\"{columname}\" t=\"str\"><x:v>{key}</x:v></x:c>");
                                        cellIndex++;
                                    }
                                }
                                writer.Write($"</x:row>");
                                yIndex++;
                            }

                            if (mode == "IDictionary<string, object>") //Dapper Row
                                GenerateSheetByDapperRow(writer, archive, value as IEnumerable, genericType, rowCount, keys.Cast<string>().ToList(), xIndex, yIndex);
                            else if (mode == "IDictionary") //IDictionary
                                GenerateSheetByIDictionary(writer, archive, value as IEnumerable, genericType, rowCount, keys, xIndex, yIndex);
                            else if (mode == "Properties")
                                GenerateSheetByProperties(writer, archive, value as IEnumerable, genericType,props, rowCount, keys, xIndex, yIndex);
                            else
                                throw new NotImplementedException($"Type {type.Name} & genericType {genericType.Name} not Implemented. please issue for me.");
                            writer.Write("</x:sheetData></x:worksheet>");
                        }
                        else if (value is DataTable)
                        {
                            GenerateSheetByDataTable(writer, archive, value as DataTable);
                        }
                        else
                        {
                            throw new NotImplementedException($"Type {type.Name} & genericType {genericType.Name} not Implemented. please issue for me.");
                        }
                        //TODO:

                    }
                End:
                    packages.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, packages);
            }
        }

        private void WriteEmptySheet(StreamWriter writer)
        {
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");
        }

        internal void GenerateSheetByDapperRow(StreamWriter writer, ZipArchive archive, IEnumerable value, Type genericType, int rowCount, List<string> keys, int xIndex = 1, int yIndex = 1)
        {
            //body
            foreach (IDictionary<string, object> v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
                foreach (var key in keys)
                {
                    var cellValue = v[key];
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
                    var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                    writer.Write($"<x:c r=\"{columname}\" {t}>");
                    writer.Write($"<x:v>{cellValueStr}");
                    writer.Write($"</x:v>");
                    writer.Write($"</x:c>");
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        internal void GenerateSheetByIDictionary(StreamWriter writer, ZipArchive archive, IEnumerable value, Type genericType, int rowCount, List<object> keys, int xIndex = 1, int yIndex = 1)
        {
            //body
            foreach (IDictionary v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
                foreach (var key in keys)
                {
                    var cellValue = v[key];
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
                    var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                    writer.Write($"<x:c r=\"{columname}\" {t}>");
                    writer.Write($"<x:v>{cellValueStr}");
                    writer.Write($"</x:v>");
                    writer.Write($"</x:c>");
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        internal void GenerateSheetByProperties(StreamWriter writer, ZipArchive archive, IEnumerable value, Type genericType, PropertyInfo[] props, int rowCount, List<object> keys, int xIndex = 1, int yIndex = 1)
        {
            //body
            foreach (var v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
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
                    var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                    writer.Write($"<x:c r=\"{columname}\" {t}>");
                    writer.Write($"<x:v>{cellValueStr}");
                    writer.Write($"</x:v>");
                    writer.Write($"</x:c>");
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        internal void GenerateSheetByDataTable(StreamWriter writer, ZipArchive archive, DataTable value)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                // dimension
                var maxRowIndex = value.Rows.Count + (printHeader && value.Rows.Count > 0 ? 1 : 0);
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

        private void GenerateContentTypesXml(ZipArchive archive, Dictionary<string, ZipPackageInfo> packages)
        {
            //[Content_Types].xml 

            var sb = new StringBuilder(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default ContentType=""application/xml"" Extension=""xml""/><Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>");
            foreach (var p in packages)
                sb.Append($"<Override ContentType=\"{p.Value.ContentType}\" PartName=\"/{p.Key}\" />");
            sb.Append("</Types>");

            ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml");
            using (var zipStream = entry.Open())
            using (StreamWriter writer = new StreamWriter(zipStream, Utf8WithBom))
                writer.Write(sb.ToString());
        }

        private string GetDimension(int maxRowIndex, int maxColumnIndex)
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

        private readonly UTF8Encoding Utf8WithBom = new System.Text.UTF8Encoding(true);
    }
}
