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
            public IDictionary<string, PropertyInfo> PropsMap { get; set; }
            public IEnumerable CellIEnumerableValues { get; set; }
        }

        private List<XRowInfo> XRowInfos { get; set; }

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
                                            xRowInfo.PropsMap = xRowInfo.IEnumerableGenricType.GetProperties().ToDictionary(s => s.Name, s => s);
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
                            var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType; //avoid nullable 
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