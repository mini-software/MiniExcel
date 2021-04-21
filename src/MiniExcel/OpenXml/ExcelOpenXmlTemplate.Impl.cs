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
using System.Text.RegularExpressions;
using System.Xml;

namespace MiniExcelLibs.OpenXml
{
    internal partial class ExcelOpenXmlTemplate
    {
        internal class XRowInfo
        {
            public string FormatText { get; set; }
            public string IEnumerablePropName { get; set; }
            public XmlElement Row { get; set; }
            public Type IEnumerableGenricType { get; set; }
            public IDictionary<string, PropInfo> PropsMap { get; set; }
            public bool IsDictionary { get; set; }
            public bool IsDataTable { get; set; }
            public IEnumerable CellIEnumerableValues { get; set; }
        }

        internal class PropInfo
        {
            public PropertyInfo PropertyInfo { get; set; }
            public Type UnderlyingTypePropType { get; set; }
        }

        private List<XRowInfo> XRowInfos { get; set; }

        private void GenerateSheetXmlImpl(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream, Dictionary<string, object> inputMaps, List<string> sharedStrings, XmlWriterSettings xmlWriterSettings = null)
        {
            var doc = new XmlDocument();
            doc.Load(sheetStream);
            sheetStream.Dispose();

            sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

            var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);
            var newSheetData = sheetData.Clone(); //avoid delete lost data
            var rows = newSheetData.SelectNodes($"x:row", _ns);

            //TODO:get merges 
            //var mergeCells = doc.SelectSingleNode($"/x:worksheet/x:mergeCells", _ns);
            //var newMergeCells = mergeCells.Clone();
            //var mergeCell = mergeCells.SelectNodes($"x:mergeCell", _ns)
            //    .Cast<XmlNode>()
            //    .ToDictionary(s=>s.Attributes["ref"].Value,s=>s);
            //foreach (var item in mergeCellsList)
            //{
            //    item.Value.RemoveAll();
            //}

            ReplaceSharedStringsToStr(sharedStrings, ref rows);
            UpdateDimensionAndGetCollectionPropertiesInfos(inputMaps, ref doc, ref rows);
            GenerateSheetXml(stream, doc, sheetData);
        }

        private void GenerateSheetXml(Stream stream, XmlDocument doc, XmlNode sheetData)
        {
            //Q.Why so complex?
            //A.Because try to use string stream avoid OOM when rendering rows
            sheetData.RemoveAll();
            sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad code smell
            var prefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $"{sheetData.Prefix}:";
            var endPrefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $":{sheetData.Prefix}"; //![image](https://user-images.githubusercontent.com/12729184/115000066-fd02b300-9ed4-11eb-8e65-bf0014015134.png)
            var contents = doc.InnerXml.Split(new string[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None); ;
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.Write(contents[0]);
                writer.Write($"<{prefix}sheetData>"); // prefix problem

                #region Generate rows and cells
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

                            if (xInfo.IsDictionary)
                            {
                                var dic = item as IDictionary<string, object>;
                                foreach (var propInfo in xInfo.PropsMap)
                                {
                                    var key = $"{{{{{xInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                    if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                    {
                                        newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                        continue;
                                    }

                                    var cellValue = dic[propInfo.Key];
                                    if (cellValue == null)
                                    {
                                        newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                        continue;
                                    }


                                    var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                                    var type = propInfo.Value.UnderlyingTypePropType;
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
                            }
                            else if (xInfo.IsDataTable)
                            {
                                var datarow = item as DataRow;
                                foreach (var propInfo in xInfo.PropsMap)
                                {
                                    var key = $"{{{{{xInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                    if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                    {
                                        newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                        continue;
                                    }

                                    var cellValue = datarow[propInfo.Key];
                                    if (cellValue == null)
                                    {
                                        newRow.InnerXml = newRow.InnerXml.Replace(key, "");
                                        continue;
                                    }


                                    var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                                    var type = propInfo.Value.UnderlyingTypePropType;
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
                            }
                            else
                            {
                                foreach (var propInfo in xInfo.PropsMap)
                                {
                                    var prop = propInfo.Value.PropertyInfo;

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
                #endregion

                writer.Write($"</{prefix}sheetData>");
                writer.Write(contents[1]);
            }
        }

        private static string CleanXml(string xml, string endPrefix)
        {
            //TODO: need to optimize
            return xml
                .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
                .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
        }

        private void ReplaceSharedStringsToStr(List<string> sharedStrings, ref XmlNodeList rows)
        {
            foreach (XmlElement row in rows)
            {
                var cs = row.SelectNodes($"x:c", _ns);
                foreach (XmlElement c in cs)
                {
                    var t = c.GetAttribute("t");
                    var v = c.SelectSingleNode("x:v", _ns);
                    if (v == null || v.InnerText == null) //![image](https://user-images.githubusercontent.com/12729184/114363496-075a3f80-9bab-11eb-9883-8e3fec10765c.png)
                        continue;

                    if (t == "s")
                    {
                        //need to check sharedstring exist or not
                        if (sharedStrings.ElementAtOrDefault(int.Parse(v.InnerText)) != null)
                        {
                            v.InnerText = sharedStrings[int.Parse(v.InnerText)];
                            // change type = str and replace its value
                            c.SetAttribute("t", "str");
                        }
                        //TODO: remove sharedstring? 
                    }
                }
            }
        }

        private void UpdateDimensionAndGetCollectionPropertiesInfos(Dictionary<string, object> inputMaps, ref XmlDocument doc, ref XmlNodeList rows)
        {
            // note : dimension need to put on the top ![image](https://user-images.githubusercontent.com/12729184/114507911-5dd88400-9c66-11eb-94c6-82ed7bdb5aab.png)

            var dimension = doc.SelectSingleNode("/x:worksheet/x:dimension", _ns) as XmlElement;
            if (dimension == null)
                throw new NotImplementedException("Excel Dimension Xml is null, please issue file for me. https://github.com/shps951023/MiniExcel/issues");
            var maxRowIndexDiff = 0;
            foreach (XmlElement row in rows)
            {
                // ==== get ienumerable infomation & maxrowindexdiff ====
                //Type ienumerableGenricType = null;
                //IDictionary<string, PropertyInfo> props = null;
                //IEnumerable ienumerable = null;
                var xRowInfo = new XRowInfo();
                xRowInfo.Row = row;
                XRowInfos.Add(xRowInfo);
                foreach (XmlElement c in row.SelectNodes($"x:c", _ns))
                {
                    var cr = c.GetAttribute("r");
                    var cLetter = new String(cr.Where(Char.IsLetter).ToArray());
                    c.SetAttribute("r", $"{cLetter}{{{{$rowindex}}}}");

                    var v = c.SelectSingleNode("x:v", _ns);
                    if (v?.InnerText == null)
                        continue;

                    var matchs = (_isExpressionRegex.Matches(v.InnerText).Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value)).ToArray();
                    var matchCnt = matchs.Length;
                    var isMultiMatch = matchCnt > 1 || (matchCnt == 1 && v.InnerText != $"{{{{{matchs[0]}}}}}");
                    foreach (var formatText in matchs)
                    {
                        xRowInfo.FormatText = formatText;
                        var propNames = formatText.Split('.');
                        if (propNames[0].StartsWith("$")) //e.g:"$rowindex" it doesn't need to check cell value type
                            continue;
                        if (!inputMaps.ContainsKey(propNames[0]))
                            throw new System.Collections.Generic.KeyNotFoundException($"Please check {propNames[0]} parameter, it's not exist.");

                        var cellValue = inputMaps[propNames[0]]; // 1. From left to right, only the first set is used as the basis for the list
                        if (cellValue is IEnumerable && !(cellValue is string))
                        {
                            xRowInfo.CellIEnumerableValues = cellValue as IEnumerable;
                            // get ienumerable runtime type
                            if (xRowInfo.IEnumerableGenricType == null) //avoid duplicate to add rowindexdiff ![image](https://user-images.githubusercontent.com/12729184/114851348-522ac000-9e14-11eb-8244-4730754d6885.png)
                            {
                                var first = true;
                                //TODO:if CellIEnumerableValues is ICollection or Array then get length or Count
                                foreach (var element in xRowInfo.CellIEnumerableValues) //TODO: optimize performance?
                                {
                                    if (xRowInfo.IEnumerableGenricType == null)
                                        if (element != null)
                                        {
                                            xRowInfo.IEnumerablePropName = propNames[0];
                                            xRowInfo.IEnumerableGenricType = element.GetType();
                                            if (element is IDictionary<string, object>)
                                            {
                                                xRowInfo.IsDictionary = true;
                                                var dic = element as IDictionary<string, object>;
                                                xRowInfo.PropsMap = dic.Keys.ToDictionary(key => key, key => dic[key] != null
                                                    ? new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(dic[key].GetType()) ?? dic[key].GetType() }
                                                    : new PropInfo { UnderlyingTypePropType = typeof(object) });
                                            }
                                            else
                                            {
                                                xRowInfo.PropsMap = xRowInfo.IEnumerableGenricType.GetProperties()
                                                    .ToDictionary(s => s.Name, s => new PropInfo { PropertyInfo = s, UnderlyingTypePropType = Nullable.GetUnderlyingType(s.PropertyType) ?? s.PropertyType });
                                            }
                                        }
                                    // ==== get demension max rowindex ====
                                    if (!first) //avoid duplicate add first one, this row not add status  ![image](https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                        maxRowIndexDiff++;
                                    first = false;
                                }
                            }

                            //TODO: check if not contain 1 index
                            //only check first one match IEnumerable, so only render one collection at same row

                            // auto check type https://github.com/shps951023/MiniExcel/issues/177
                            var prop = xRowInfo.PropsMap[propNames[1]];
                            var type = prop.UnderlyingTypePropType; //avoid nullable 
                                                                    // 
                            if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                                throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");

                            if (isMultiMatch)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (Helpers.IsNumericType(type))
                            {
                                c.SetAttribute("t", "n");
                            }
                            else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                            {
                                c.SetAttribute("t", "b");
                            }
                            else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                            {
                                c.SetAttribute("t", "str");
                            }

                            break;
                        }
                        else if (cellValue is DataTable)
                        {
                            var dt = cellValue as DataTable;
                            if (xRowInfo.CellIEnumerableValues == null)
                            {
                                xRowInfo.IEnumerablePropName = propNames[0];
                                xRowInfo.IEnumerableGenricType = typeof(DataRow);
                                xRowInfo.IsDataTable = true;
                                xRowInfo.CellIEnumerableValues = dt.Rows.Cast<DataRow>(); //TODO: need to optimize performance
                                var first = true;
                                foreach (var element in xRowInfo.CellIEnumerableValues)
                                {
                                    // ==== get demension max rowindex ====
                                    if (!first) //avoid duplicate add first one, this row not add status  ![image](https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                        maxRowIndexDiff++;
                                    first = false;
                                }
                                //TODO:need to optimize
                                //maxRowIndexDiff = dt.Rows.Count <= 1 ? 0 : dt.Rows.Count-1;
                                xRowInfo.PropsMap = dt.Columns.Cast<DataColumn>().ToDictionary(col => col.ColumnName, col =>
                                new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(col.DataType) }
                            );
                            }

                            var column = dt.Columns[propNames[1]];
                            var type = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType; //avoid nullable 
                            if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                                throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");

                            if (isMultiMatch)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (Helpers.IsNumericType(type))
                            {
                                c.SetAttribute("t", "n");
                            }
                            else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                            {
                                c.SetAttribute("t", "b");
                            }
                            else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                            {
                                c.SetAttribute("t", "str");
                            }
                        }
                        else
                        {
                            var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue);
                            if (isMultiMatch) // if matchs count over 1 need to set type=str ![image](https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (decimal.TryParse(cellValueStr, out var outV))
                            {
                                c.SetAttribute("t", "n");
                            }
                            else if (cellValue is bool)
                            {
                                c.SetAttribute("t", "b");
                                cellValueStr = (bool)cellValue ? "1" : "0";
                            }
                            else if (cellValue is DateTime || cellValue is DateTime?)
                            {
                                //c.SetAttribute("t", "d");
                                cellValueStr = ((DateTime)cellValue).ToString("yyyy-MM-dd HH:mm:ss");
                            }

                            v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", cellValueStr); //TODO: auto check type and set value
                        }
                    }
                    //if (xRowInfo.CellIEnumerableValues != null) //2. From left to right, only the first set is used as the basis for the list
                    //    break;
                }
            }

            // e.g <dimension ref=\"A1:B6\" /> only need to update B6 to BMaxRowIndex
            var @refs = dimension.GetAttribute("ref").Split(':');
            var letter = new String(refs[1].Where(Char.IsLetter).ToArray());
            var digit = int.Parse(new String(refs[1].Where(Char.IsDigit).ToArray()));

            dimension.SetAttribute("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
        }
    }
}