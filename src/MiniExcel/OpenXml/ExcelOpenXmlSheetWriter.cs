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
using static MiniExcelLibs.Utils.Helpers;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlSheetWriter : IExcelWriter
    {
        private readonly static UTF8Encoding _utf8WithBom = new System.Text.UTF8Encoding(true);
        private Stream _stream;
        public ExcelOpenXmlSheetWriter(Stream stream)
        {
            this._stream = stream;
        }

        public void SaveAs(object value, string sheetName, bool printHeader, IConfiguration configuration)
        {
            using (var archive = new MiniExcelZipArchive(_stream, ZipArchiveMode.Create, true, _utf8WithBom))
            {
                var packages = DefualtOpenXml.GenerateDefaultOpenXml(archive, sheetName);
                var sheetPath = "xl/worksheets/sheet1.xml";
                {
                    ZipArchiveEntry entry = archive.CreateEntry(sheetPath);
                    using (var zipStream = entry.Open())
                    using (StreamWriter writer = new StreamWriter(zipStream, _utf8WithBom))
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
                            List<ExcelCustomPropertyInfo> props = null;
                            string mode = null;

                            // reason : https://stackoverflow.com/questions/66797421/how-replace-top-format-mark-after-streamwriter-writing
                            // check mode & get maxRowCount & maxColumnIndex
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
                                            if (genericType.IsValueType)
                                                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} value generic type");
                                            else if (genericType == typeof(string) || genericType == typeof(DateTime) || genericType == typeof(Guid))
                                                throw new NotImplementedException($"MiniExcel not support only {genericType.Name} generic type");
                                            props = Helpers.GetSaveAsProperties(genericType);
                                            maxColumnIndex = props.Count;
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
                            writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

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
                                        if (p == null)
                                        {
                                            cellIndex++; //reason : https://github.com/shps951023/MiniExcel/issues/142
                                            continue;
                                        }

                                        var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
                                        writer.Write($"<x:c r=\"{columname}\" t=\"str\"><x:v>{p.ExcelColumnName}</x:v></x:c>");
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

                            // body
                            if (mode == "IDictionary<string, object>") //Dapper Row
                                GenerateSheetByDapperRow(writer, archive, value as IEnumerable, rowCount, keys.Cast<string>().ToList(), xIndex, yIndex);
                            else if (mode == "IDictionary") //IDictionary
                                GenerateSheetByIDictionary(writer, archive, value as IEnumerable, rowCount, keys, xIndex, yIndex);
                            else if (mode == "Properties")
                                GenerateSheetByProperties(writer, archive, value as IEnumerable, props, rowCount, xIndex, yIndex);
                            else
                                throw new NotImplementedException($"Type {type.Name} & genericType {genericType.Name} not Implemented. please issue for me.");
                            writer.Write("</x:sheetData></x:worksheet>");
                        }
                        else if (value is DataTable)
                        {
                            GenerateSheetByDataTable(writer, archive, value as DataTable, printHeader);
                        }
                        else
                        {
                            throw new NotImplementedException($"Type {type.Name} & genericType {genericType.Name} not Implemented. please issue for me.");
                        }
                    }
                End:
                    packages.Add(sheetPath, new ZipPackageInfo(entry, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                }
                GenerateContentTypesXml(archive, packages);
            }
        }

        private void WriteEmptySheet(StreamWriter writer)
        {
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><x:dimension ref=""A1""/><x:sheetData></x:sheetData></x:worksheet>");
        }

        private void GenerateSheetByDapperRow(StreamWriter writer, MiniExcelZipArchive archive, IEnumerable value, int rowCount, List<string> keys, int xIndex = 1, int yIndex = 1)
        {
            foreach (IDictionary<string, object> v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
                foreach (var key in keys)
                {
                    var cellValue = v[key];
                    WriteCell(writer, yIndex, cellIndex, cellValue);
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        private void GenerateSheetByIDictionary(StreamWriter writer, MiniExcelZipArchive archive, IEnumerable value, int rowCount, List<object> keys, int xIndex = 1, int yIndex = 1)
        {
            foreach (IDictionary v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
                foreach (var key in keys)
                {
                    var cellValue = v[key];
                    WriteCell(writer, yIndex, cellIndex, cellValue);
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        private void GenerateSheetByProperties(StreamWriter writer, MiniExcelZipArchive archive, IEnumerable value, List<ExcelCustomPropertyInfo> props, int rowCount, int xIndex = 1, int yIndex = 1)
        {
            foreach (var v in value)
            {
                writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                var cellIndex = xIndex;
                foreach (var p in props)
                {
                    if (p == null) //reason:https://github.com/shps951023/MiniExcel/issues/142
                    {
                        cellIndex++;
                        continue;
                    }
                    var cellValue = p.Property.GetValue(v);
                    WriteCell(writer, yIndex, cellIndex, cellValue);
                    cellIndex++;
                }
                writer.Write($"</x:row>");
                yIndex++;
            }
        }

        private static void WriteCell(StreamWriter writer, int yIndex, int cellIndex, object value)
        {
            var v = string.Empty;
            var t = string.Empty;
            if (value == null)
            {
                t = "t=\"str\"";
                v = "";
            }
            else if (value is string)
            {
                t = "t=\"str\"";
                v = ExcelOpenXmlUtils.EncodeXML(value.ToString());
            }
            else
            {
                var type = value.GetType();
                type = Nullable.GetUnderlyingType(type) ?? type;

                if (Helpers.IsNumericType(type))
                {
                    t = "t=\"n\"";
                    v = value.ToString();
                }
                else if (type == typeof(bool))
                {
                    t = "t=\"b\"";
                    v = (bool)value ? "1" : "0";
                }
                else if (type == typeof(DateTime))
                {
                    t = "s=\"1\"";
                    v = ((DateTime)value).ToOADate().ToString();
                }
                else
                {
                    t = "t=\"str\"";
                    v = ExcelOpenXmlUtils.EncodeXML(value.ToString());
                }
            }

            var columname = ExcelOpenXmlUtils.ConvertXyToCell(cellIndex, yIndex);
            writer.Write($"<x:c r=\"{columname}\" {t}><x:v>{v}</x:v></x:c>");
        }

        private void GenerateSheetByDataTable(StreamWriter writer, MiniExcelZipArchive archive, DataTable value, bool printHeader)
        {
            var xy = ExcelOpenXmlUtils.ConvertCellToXY("A1");

            //GOTO Top Write:
            writer.Write($@"<?xml version=""1.0"" encoding=""utf-8""?><x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            {
                var yIndex = xy.Item2;

                // dimension
                var maxRowIndex = value.Rows.Count + (printHeader && value.Rows.Count > 0 ? 1 : 0);
                var maxColumnIndex = value.Columns.Count;
                writer.Write($@"<x:dimension ref=""{GetDimensionRef(maxRowIndex, maxColumnIndex)}""/><x:sheetData>");

                if (printHeader)
                {
                    writer.Write($"<x:row r=\"{yIndex.ToString()}\">");
                    var xIndex = xy.Item1;
                    foreach (DataColumn c in value.Columns)
                    {
                        var r = ExcelOpenXmlUtils.ConvertXyToCell(xIndex, yIndex);
                        writer.Write($"<x:c r=\"{r}\" t=\"str\">");
                        writer.Write($"<x:v>{c.Caption ?? c.ColumnName}");
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
                        WriteCell(writer, yIndex, xIndex, cellValue);
                        xIndex++;
                    }
                    writer.Write($"</x:row>");
                    yIndex++;
                }
            }
            writer.Write("</x:sheetData></x:worksheet>");
        }

        private void GenerateContentTypesXml(MiniExcelZipArchive archive, Dictionary<string, ZipPackageInfo> packages)
        {
            //[Content_Types].xml 

            var sb = new StringBuilder(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types""><Default ContentType=""application/xml"" Extension=""xml""/><Default ContentType=""application/vnd.openxmlformats-package.relationships+xml"" Extension=""rels""/>");
            foreach (var p in packages)
                sb.Append($"<Override ContentType=\"{p.Value.ContentType}\" PartName=\"/{p.Key}\" />");
            sb.Append("</Types>");

            ZipArchiveEntry entry = archive.CreateEntry("[Content_Types].xml");
            using (var zipStream = entry.Open())
            using (StreamWriter writer = new StreamWriter(zipStream, _utf8WithBom))
                writer.Write(sb.ToString());
        }

        private string GetDimensionRef(int maxRowIndex, int maxColumnIndex)
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
    }
}
