
namespace MiniExcelLibs.OpenXml
{
    using MiniExcelLibs.Utils;
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

    internal partial class ExcelOpenXmlTemplate:IExcelTemplate
    {
        private static readonly XmlNamespaceManager _ns;
        private static readonly Regex _isExpressionRegex;
        static ExcelOpenXmlTemplate()
        {
            _isExpressionRegex = new Regex("(?<={{).*?(?=}})");
            _ns = new XmlNamespaceManager(new NameTable());
            _ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        }

        private readonly Stream stream;
        public ExcelOpenXmlTemplate(Stream _strem)
        {
            stream = _strem;
        }

        public void SaveAsByTemplate(string templatePath, object value)
        {
            using (var stream = File.Open(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                SaveAsByTemplateImpl(stream, value);
        }
        public void SaveAsByTemplate(byte[] templateBtyes, object value)
        {
            using (Stream stream = new MemoryStream(templateBtyes))
                SaveAsByTemplateImpl(stream, value);
        }

        public void SaveAsByTemplateImpl(Stream templateStream, object value)
        {
            //only support xlsx         
            Dictionary<string, object> values = null;
            if (value is Dictionary<string, object>)
            {
                values = value as Dictionary<string, object>;
            }
            else
            {
                var type = value.GetType();
                var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
                values = new Dictionary<string, object>();
                foreach (var p in props)
                {
                    values.Add(p.Name, p.GetValue(value));
                }
            }
            //TODO:DataTable & DapperRow
            {
                templateStream.CopyTo(stream);

                var reader = new ExcelOpenXmlSheetReader(stream);
                var _archive = new ExcelOpenXmlZip(stream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);
                {
                    //read sharedString
                    var sharedStrings = reader.GetSharedStrings();

                    //read all xlsx sheets
                    var sheets = _archive.ZipFile.Entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                        || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    ).ToList();

                    foreach (var sheet in sheets)
                    {
                        this.XRowInfos = new List<XRowInfo>(); //every time need to use new XRowInfos or it'll cause duplicate problem: https://user-images.githubusercontent.com/12729184/115003101-0fcab700-9ed8-11eb-9151-ca4d7b86d59e.png
                        var sheetStream = sheet.Open();
                        var fullName = sheet.FullName;

                        ZipArchiveEntry entry = _archive.ZipFile.CreateEntry(fullName);
                        using (var zipStream = entry.Open())
                        {
                            GenerateSheetXmlImpl(sheet, zipStream, sheetStream, values, sharedStrings);
                            //doc.Save(zipStream); //don't do it beacause : ![image](https://user-images.githubusercontent.com/12729184/114361127-61a5d100-9ba8-11eb-9bb9-34f076ee28a2.png)
                        }
                    }
                }

                _archive.ZipFile.Dispose();
            }
        }

        private void GenerateSheetXmlImpl(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream, Dictionary<string, object> inputMaps, List<string> sharedStrings, XmlWriterSettings xmlWriterSettings = null)
        {
            var doc = new XmlDocument();
            doc.Load(sheetStream);
            sheetStream.Dispose();

            sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

            var worksheet = doc.SelectSingleNode("/x:worksheet", _ns);
            var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);

            var newSheetData = sheetData.Clone(); //avoid delete lost data
            var rows = newSheetData.SelectNodes($"x:row", _ns);

            ReplaceSharedStringsToStr(sharedStrings, ref rows);

            //Update dimension && Check if the column contains a collection and get type and properties infomations
            UpdateDimensionAndGetCollectionPropertiesInfos(inputMaps, ref doc, ref rows);

            #region Render cell values

            //Q.Why so complex?
            //A.Because try to use string stream avoid OOM when rendering rows
            sheetData.RemoveAll();
            sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad smell
            var prefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $"{sheetData.Prefix}:";
            var endPrefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $":{sheetData.Prefix}"; //![image](https://user-images.githubusercontent.com/12729184/115000066-fd02b300-9ed4-11eb-8e65-bf0014015134.png)
            var contents = doc.InnerXml.Split(new string[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None); ;
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.Write(contents[0]);
                writer.Write($"<{prefix}sheetData>"); // prefix problem
                int originRowIndex;
                int rowIndexDiff = 0;
                foreach (var xInfo in XRowInfos)
                {
                    var row = xInfo.Row;

                    //TODO: some xlsx without r
                    originRowIndex = int.Parse(row.GetAttribute("r"));
                    var newRowIndex = originRowIndex + rowIndexDiff;
                    if (xInfo.CellIEnumerableValues != null)
                    {
                        var first = true;
                        foreach (var item in xInfo.CellIEnumerableValues)
                        {
                            var newRow = row.Clone() as XmlElement;
                            newRow.SetAttribute("r", newRowIndex.ToString());
                            newRow.InnerXml = row.InnerXml.Replace($"{{{{$rowindex}}}}", newRowIndex.ToString());

                            foreach (var propInfo in xInfo.PropsMap)
                            {
                                var prop = propInfo.Value;

                                var key = $"{{{{{xInfo.IEnumerablePropName}.{prop.Name}}}}}";
                                if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                {
                                    newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                    continue;
                                }

                                var cellValue = prop.GetValue(item);
                                if (cellValue == null)
                                {
                                    newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                    continue;
                                }
                                    

                                var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                                var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                if (type == typeof(bool))
                                {
                                    cellValueStr = (bool)cellValue ? "1" : "0";
                                }
                                else if (type == typeof(DateTime))
                                {
                                    //c.SetAttribute("t", "d");
                                    cellValueStr = ((DateTime)cellValue).ToString("yyyy-MM-dd HH:mm:ss");
                                }

                                //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)
                                newRow.InnerXml = newRow.InnerXml.Replace(key, cellValueStr);
                            }

                            // note: only first time need add diff ![image](https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png)
                            if (!first)
                                rowIndexDiff++;
                            first = false;

                            newRowIndex++;
                            writer.Write(CleanXml(newRow.OuterXml, endPrefix));
                            newRow = null;
                        }
                    }
                    else
                    {
                        row.SetAttribute("r", newRowIndex.ToString());
                        row.InnerXml = row.InnerXml.Replace($"{{{{$rowindex}}}}", newRowIndex.ToString());
                        writer.Write(CleanXml(row.OuterXml, endPrefix));
                    }
                }
                writer.Write($"</{prefix}sheetData>");
                writer.Write(contents[1]);
            }
            #endregion
        }

        private static string CleanXml(string xml,string endPrefix)
        {
            //TODO: need to optimize
            return xml
                .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
                .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
        }
    }
}
