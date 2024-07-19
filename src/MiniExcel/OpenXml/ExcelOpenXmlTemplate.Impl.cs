using MiniExcelLibs.Attributes;
using MiniExcelLibs.Utils;
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
        public class XRowInfo
        {
            public string FormatText { get; set; }
            public string IEnumerablePropName { get; set; }
            public XmlElement Row { get; set; }
            public Type IEnumerableGenricType { get; set; }
            public IDictionary<string, PropInfo> PropsMap { get; set; }
            public bool IsDictionary { get; set; }
            public bool IsDataTable { get; set; }
            public int CellIEnumerableValuesCount { get; set; }
            public IList<object> CellIlListValues { get; set; }
            public IEnumerable CellIEnumerableValues { get; set; }
            public XMergeCell IEnumerableMercell { get; set; }
            public List<XMergeCell> RowMercells { get; set; }
        }

        public class PropInfo
        {
            public PropertyInfo PropertyInfo { get; set; }
            public FieldInfo FieldInfo { get; set; }
            public Type UnderlyingTypePropType { get; set; }
            public PropertyInfoOrFieldInfo PropertyInfoOrFieldInfo { get; set; } = PropertyInfoOrFieldInfo.None;
        }

        public enum PropertyInfoOrFieldInfo
        {
            None = 0,
            PropertyInfo = 1,
            FieldInfo = 2
        }

        public class XMergeCell
        {
            public XMergeCell(XMergeCell mergeCell)
            {
                this.Width = mergeCell.Width;
                this.Height = mergeCell.Height;
                this.X1 = mergeCell.X1;
                this.Y1 = mergeCell.Y1;
                this.X2 = mergeCell.X2;
                this.Y2 = mergeCell.Y2;
                this.MergeCell = mergeCell.MergeCell;
            }
            public XMergeCell(XmlElement mergeCell)
            {
                var @ref = mergeCell.Attributes["ref"].Value;
                var refs = @ref.Split(':');

                //TODO: width,height
                var xy1 = refs[0];
                X1 = ColumnHelper.GetColumnIndex(StringHelper.GetLetter(refs[0]));
                Y1 = StringHelper.GetNumber(xy1);

                var xy2 = refs[1];
                X2 = ColumnHelper.GetColumnIndex(StringHelper.GetLetter(refs[1]));
                Y2 = StringHelper.GetNumber(xy2);

                Width = Math.Abs(X1 - X2) + 1;
                Height = Math.Abs(Y1 - Y2) + 1;
            }
            public XMergeCell(string x1, int y1, string x2, int y2)
            {
                X1 = ColumnHelper.GetColumnIndex(x1);
                Y1 = y1;

                X2 = ColumnHelper.GetColumnIndex(x2);
                Y2 = y2;

                Width = Math.Abs(X1 - X2) + 1;
                Height = Math.Abs(Y1 - Y2) + 1;
            }

            public string XY1 { get { return $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}"; } }
            public int X1 { get; set; }
            public int Y1 { get; set; }
            public string XY2 { get { return $"{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}"; } }
            public int X2 { get; set; }
            public int Y2 { get; set; }
            public string Ref { get { return $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}"; } }
            public XmlElement MergeCell { get; set; }
            public int Width { get; internal set; }
            public int Height { get; internal set; }

            public string ToXmlString(string prefix)
            {
                return $"<{prefix}mergeCell ref=\"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}\"/>";
            }
        }

		private List<XRowInfo> XRowInfos { get; set; }

        private readonly List<string> CalcChainCellRefs = new List<string>();

        private Dictionary<string, XMergeCell> XMergeCellInfos { get; set; }
        public List<XMergeCell> NewXMergeCellInfos { get; private set; }

        private void GenerateSheetXmlImpl(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream,
            Dictionary<string, object> inputMaps, IDictionary<int, string> sharedStrings,
            bool mergeCells = false)
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
            GetMercells(doc, worksheet);
            UpdateDimensionAndGetRowsInfo(inputMaps, ref doc, ref rows, !mergeCells);

            WriteSheetXml(stream, doc, sheetData, mergeCells);

        }

        private void GetMercells(XmlDocument doc, XmlNode worksheet)
        {
            var mergeCells = doc.SelectSingleNode($"/x:worksheet/x:mergeCells", _ns);
            if (mergeCells != null)
            {
                var newMergeCells = mergeCells.Clone();
                //mergeCells.RemoveAll();
                worksheet.RemoveChild(mergeCells);

                foreach (XmlElement cell in newMergeCells)
                {
                    var @ref = cell.Attributes["ref"].Value;
                    var mergerCell = new XMergeCell(cell);
                    this.XMergeCellInfos[mergerCell.XY1] = mergerCell;
                }
            }
        }

        private class MergeCellIndex
        {
            public int RowStart { get; set; }
            public int RowEnd { get; set; }

            public MergeCellIndex(int rowStart, int rowEnd)
            {
                RowStart = rowStart;
                RowEnd = rowEnd;
            }
        }

        private class XChildNode
        {
            public string InnerText { get; set; }
            public string ColIndex { get; set; }
            public int RowIndex { get; set; }
        }

        private void WriteSheetXml(Stream stream, XmlDocument doc, XmlNode sheetData, bool mergeCells = false)
        {
            //Q.Why so complex?
            //A.Because try to use string stream avoid OOM when rendering rows
            sheetData.RemoveAll();
            sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad code smell
            var prefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $"{sheetData.Prefix}:";
            var endPrefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $":{sheetData.Prefix}"; //![image](https://user-images.githubusercontent.com/12729184/115000066-fd02b300-9ed4-11eb-8e65-bf0014015134.png)
            var contents = doc.InnerXml.Split(new string[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None);
            using (var writer = new StreamWriter(stream, Encoding.UTF8))
            {
                writer.Write(contents[0]);
                writer.Write($"<{prefix}sheetData>"); // prefix problem

                #region MergeCells

                if (mergeCells)
                {
                    var mergeTaggedColumns = new Dictionary<XChildNode, XChildNode>();
                    var columns = XRowInfos.SelectMany(s => s.Row.Cast<XmlElement>())
                        .Where(s => !string.IsNullOrEmpty(s.InnerText)).Select(s =>
                        {
                            var att = s.GetAttribute("r");
                            return new XChildNode()
                            {
                                InnerText = s.InnerText,
                                ColIndex = StringHelper.GetLetter(att),
                                RowIndex = StringHelper.GetNumber(att)
                            };
                        }).OrderBy(x => x.RowIndex).ToList();

                    var mergeColumns = columns.Where(s => s.InnerText.Contains("@merge")).ToList();
                    var endMergeColumns = columns.Where(s => s.InnerText.Contains("@endmerge")).ToList();
                    var mergeLimitColumn = mergeColumns.FirstOrDefault(x => x.InnerText.Contains("@mergelimit"));

                    foreach (var mergeColumn in mergeColumns)
                    {
                        var endMergeColumn = endMergeColumns.FirstOrDefault(s =>
                            s.ColIndex == mergeColumn.ColIndex && s.RowIndex > mergeColumn.RowIndex);

                        if (endMergeColumn != null)
                        {
                            mergeTaggedColumns[mergeColumn] = endMergeColumn;
                        }
                    }

                    var calculatedColumns = new List<XChildNode>();

                    if (mergeTaggedColumns.Count > 0)
                    {
                        foreach (var taggedColumn in mergeTaggedColumns)
                        {
                            calculatedColumns.AddRange(columns.Where(x =>
                                x.ColIndex == taggedColumn.Key.ColIndex && x.RowIndex > taggedColumn.Key.RowIndex &&
                                x.RowIndex < taggedColumn.Value.RowIndex));
                        }

                        Dictionary<int, MergeCellIndex>
                            lastMergeCellIndexes = new Dictionary<int, MergeCellIndex>();

                        for (int rowNo = 0; rowNo < XRowInfos.Count; rowNo++)
                        {
                            var rowInfo = XRowInfos[rowNo];
                            var row = rowInfo.Row;
                            var childNodes = row.ChildNodes.Cast<XmlElement>().ToList();

                            foreach (var childNode in childNodes)
                            {
                                var att = childNode.GetAttribute("r");
                                var childNodeLetter = StringHelper.GetLetter(att);
                                var childNodeNumber = StringHelper.GetNumber(att);

                                if (!string.IsNullOrEmpty(childNode.InnerText))
                                {
                                    var xmlNodes = calculatedColumns
                                        .Where(j => j.InnerText == childNode.InnerText && j.ColIndex == childNodeLetter)
                                        .OrderBy(s => s.RowIndex).ToList();

                                    if (xmlNodes.Count > 1)
                                    {
                                        if (mergeLimitColumn != null)
                                        {
                                            var limitedNode = calculatedColumns.First(j =>
                                                j.ColIndex == mergeLimitColumn.ColIndex && j.RowIndex == childNodeNumber);

                                            var limitedMaxNode = calculatedColumns.Last(j =>
                                                j.ColIndex == mergeLimitColumn.ColIndex && j.InnerText == limitedNode.InnerText);

                                            xmlNodes = xmlNodes.Where(j => j.RowIndex >= limitedNode.RowIndex && j.RowIndex <= limitedMaxNode.RowIndex).ToList();
                                        }

                                        var firstRow = xmlNodes.FirstOrDefault();
                                        var lastRow = xmlNodes.LastOrDefault(s =>
                                            s.RowIndex <= firstRow?.RowIndex + xmlNodes.Count &&
                                            s.RowIndex != firstRow?.RowIndex);

                                        if (firstRow != null && lastRow != null)
                                        {
                                            var mergeCell = new XMergeCell(firstRow.ColIndex, firstRow.RowIndex,
                                                lastRow.ColIndex, lastRow.RowIndex);

                                            var mergeIndexResult =
                                                lastMergeCellIndexes.TryGetValue(mergeCell.X1, out var mergeIndex);

                                            if (!mergeIndexResult || mergeCell.Y1 < mergeIndex.RowStart ||
                                                mergeCell.Y2 > mergeIndex.RowEnd)
                                            {
                                                lastMergeCellIndexes[mergeCell.X1] =
                                                    new MergeCellIndex(mergeCell.Y1, mergeCell.Y2);

                                                if (rowInfo.RowMercells == null)
                                                {
                                                    rowInfo.RowMercells = new List<XMergeCell>();
                                                }

                                                rowInfo.RowMercells.Add(mergeCell);
                                            }
                                        }
                                    }
                                }

                                childNode.SetAttribute("r", $"{childNodeLetter}{{{{$rowindex}}}}"); //TODO:
                            }
                        }
                    }
                }

                #endregion

                #region Generate rows and cells
                int originRowIndex;
                int rowIndexDiff = 0;
                var rowXml = new StringBuilder();

                // for formula cells
                int enumrowstart = -1;
                int enumrowend = -1;

                // for grouped cells
                bool groupingStarted = false;
                bool hasEverGroupStarted = false;
                int groupStartRowIndex = 0;
                IList<object> cellIEnumerableValues = null;
                bool isCellIEnumerableValuesSet = false;
                int cellIEnumerableValuesIndex = 0;
                int groupRowCount = 0;
                int headerDiff = 0;
                bool isFirstRound = true;
                string currentHeader = "";
                string prevHeader = "";
                bool isHeaderRow = false;
                int mergeRowCount = 0;

                for (int rowNo = 0; rowNo < XRowInfos.Count; rowNo++)
                {
                    isHeaderRow = false;
                    currentHeader = "";

                    var rowInfo = XRowInfos[rowNo];
                    var row = rowInfo.Row;

                    if (row.InnerText.Contains("@group"))
                    {
                        groupingStarted = true;
                        hasEverGroupStarted = true;
                        groupStartRowIndex = rowNo;
                        isFirstRound = true;
                        prevHeader = "";
                        continue;
                    }
                    else if (row.InnerText.Contains("@endgroup"))
                    {
                        if (cellIEnumerableValuesIndex >= cellIEnumerableValues.Count - 1)
                        {
                            groupingStarted = false;
                            groupStartRowIndex = 0;
                            cellIEnumerableValues = null;
                            isCellIEnumerableValuesSet = false;
                            headerDiff++;
                            continue;
                        }
                        rowNo = groupStartRowIndex;
                        cellIEnumerableValuesIndex++;
                        isFirstRound = false;
                        continue;
                    }
                    else if (row.InnerText.Contains("@header"))
                    {
                        isHeaderRow = true;
                    }
                    else if (row.InnerText.Contains("@merge") && mergeCells)
                    {
                        mergeRowCount++;
                        continue;
                    }
                    else if (row.InnerText.Contains("@endmerge") && mergeCells)
                    {
                        mergeRowCount++;
                        continue;
                    }

                    if (groupingStarted && !isCellIEnumerableValuesSet && rowInfo.CellIlListValues != null)
                    {
                        cellIEnumerableValues = rowInfo.CellIlListValues;
                        isCellIEnumerableValuesSet = true;
                    }

                    var groupingRowDiff =
                        (hasEverGroupStarted ? (-1 + cellIEnumerableValuesIndex * groupRowCount - headerDiff) : 0);

                    if (groupingStarted)
                    {
                        if (isFirstRound)
                        {
                            groupRowCount++;
                        }

                        if (cellIEnumerableValues != null)
                        {
                            rowInfo.CellIEnumerableValuesCount = 1;
                            rowInfo.CellIEnumerableValues =
                                cellIEnumerableValues.Skip(cellIEnumerableValuesIndex).Take(1).ToList();
                        }
                    }

                    //TODO: some xlsx without r
                    originRowIndex = int.Parse(row.GetAttribute("r"));
                    var newRowIndex = originRowIndex + rowIndexDiff + groupingRowDiff - mergeRowCount;

                    string innerXml = row.InnerXml;
                    rowXml.Clear()
                          .AppendFormat(@"<{0}", row.Name);
                    foreach (var attr in row.Attributes.Cast<XmlAttribute>()
                                                       .Where(e => e.Name != "r"))
                    {
                        rowXml.AppendFormat(@" {0}=""{1}""", attr.Name, attr.Value);
                    }

                    string outerXmlOpen = rowXml.ToString();

                    if (rowInfo.CellIEnumerableValues != null)
                    {
                        var first = true;
                        var iEnumerableIndex = 0;
                        enumrowstart = newRowIndex;

                        foreach (var item in rowInfo.CellIEnumerableValues)
                        {
                            iEnumerableIndex++;

                            rowXml.Clear()
                                  .Append(outerXmlOpen)
                                  .AppendFormat(@" r=""{0}"">", newRowIndex)
                                  .Append(innerXml)
                                  .Replace($"{{{{$rowindex}}}}", newRowIndex.ToString())
                                  .AppendFormat(@"</{0}>", row.Name);

                            var rowXmlString = rowXml.ToString();
                            var extract = "";
                            var newCellValue = "";

                            var ifIndex = rowXmlString.IndexOf("@if", StringComparison.Ordinal);
                            var endifIndex = rowXmlString.IndexOf("@endif", StringComparison.Ordinal);

                            if (ifIndex != -1 && endifIndex != -1)
                            {
                                extract = rowXmlString.Substring(ifIndex, endifIndex - ifIndex + 6);
                            }

                            var lines = extract.Split('\n');

                            if (rowInfo.IsDictionary)
                            {
                                var dic = item as IDictionary<string, object>;

                                for (var i = 0; i < lines.Length; i++)
                                {
                                    if (lines[i].Contains("@if") || lines[i].Contains("@elseif"))
                                    {
                                        var newLines = lines[i].Replace("@elseif(", "").Replace("@if(", "").Replace(")", "").Split(' ');

                                        var value = dic[newLines[0]];
                                        var evaluation = EvaluateStatement(value, newLines[1], newLines[2]);

                                        if (evaluation)
                                        {
                                            newCellValue += lines[i + 1];
                                            break;
                                        }
                                    }
                                    else if (lines[i].Contains("@else"))
                                    {
                                        newCellValue += lines[i + 1];
                                        break;
                                    }
                                }

                                if (!string.IsNullOrEmpty(newCellValue))
                                {
                                    rowXml.Replace(extract, newCellValue);
                                }

                                foreach (var propInfo in rowInfo.PropsMap)
                                {
                                    var key = $"{{{{{rowInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                    if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }
                                    if (!dic.ContainsKey(propInfo.Key))
                                        continue;
                                    var cellValue = dic[propInfo.Key];
                                    if (cellValue == null)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }


                                    var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                    var type = propInfo.Value.UnderlyingTypePropType;
                                    if (type == typeof(bool))
                                    {
                                        cellValueStr = (bool)cellValue ? "1" : "0";
                                    }
                                    else if (type == typeof(DateTime))
                                    {
                                        cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                    }

                                    //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)

                                    rowXml.Replace("@header" + key, cellValueStr);
                                    rowXml.Replace(key, cellValueStr);

                                    if (isHeaderRow && row.InnerText.Contains(key))
                                    {
                                        currentHeader += cellValueStr;
                                    }
                                }
                            }
                            else if (rowInfo.IsDataTable)
                            {
                                var datarow = item as DataRow;

                                for (var i = 0; i < lines.Length; i++)
                                {
                                    if (lines[i].Contains("@if") || lines[i].Contains("@elseif"))
                                    {
                                        var newLines = lines[i].Replace("@elseif(", "").Replace("@if(", "").Replace(")", "").Split(' ');

                                        var value = datarow[newLines[0]];
                                        var evaluation = EvaluateStatement(value, newLines[1], newLines[2]);

                                        if (evaluation)
                                        {
                                            newCellValue += lines[i + 1];
                                            break;
                                        }
                                    }
                                    else if (lines[i].Contains("@else"))
                                    {
                                        newCellValue += lines[i + 1];
                                        break;
                                    }
                                }

                                if (!string.IsNullOrEmpty(newCellValue))
                                {
                                    rowXml.Replace(extract, newCellValue);
                                }

                                foreach (var propInfo in rowInfo.PropsMap)
                                {
                                    var key = $"{{{{{rowInfo.IEnumerablePropName}.{propInfo.Key}}}}}";
                                    if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }

                                    var cellValue = datarow[propInfo.Key];
                                    if (cellValue == null)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }


                                    var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                    var type = propInfo.Value.UnderlyingTypePropType;
                                    if (type == typeof(bool))
                                    {
                                        cellValueStr = (bool)cellValue ? "1" : "0";
                                    }
                                    else if (type == typeof(DateTime))
                                    {
                                        cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                    }

                                    //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)

                                    rowXml.Replace("@header" + key, cellValueStr);
                                    rowXml.Replace(key, cellValueStr);

                                    if (isHeaderRow && row.InnerText.Contains(key))
                                    {
                                        currentHeader += cellValueStr;
                                    }
                                }
                            }
                            else
                            {
                                for (var i = 0; i < lines.Length; i++)
                                {
                                    if (lines[i].Contains("@if") || lines[i].Contains("@elseif"))
                                    {
                                        var newLines = lines[i].Replace("@elseif(", "").Replace("@if(", "").Replace(")", "").Split(' ');

                                        var prop = rowInfo.PropsMap[newLines[0]];
                                        object value = string.Empty;
                                        if (prop.PropertyInfoOrFieldInfo == PropertyInfoOrFieldInfo.PropertyInfo)
                                        {
                                            value = rowInfo.PropsMap[newLines[0]].PropertyInfo.GetValue(item);
                                        }
                                        else if (prop.PropertyInfoOrFieldInfo == PropertyInfoOrFieldInfo.FieldInfo)
                                        {
                                            value = rowInfo.PropsMap[newLines[0]].FieldInfo.GetValue(item);
                                        }

                                        var evaluation = EvaluateStatement(value, newLines[1], newLines[2]);

                                        if (evaluation)
                                        {
                                            newCellValue += lines[i + 1];
                                            break;
                                        }
                                    }
                                    else if (lines[i].Contains("@else"))
                                    {
                                        newCellValue += lines[i + 1];
                                        break;
                                    }
                                }

                                if (!string.IsNullOrEmpty(newCellValue))
                                {
                                    rowXml.Replace(extract, newCellValue);
                                }

                                foreach (var propInfo in rowInfo.PropsMap)
                                {
                                    var prop = propInfo.Value.PropertyInfo;

                                    var key = $"{{{{{rowInfo.IEnumerablePropName}.{prop.Name}}}}}";
                                    if (item == null) //![image](https://user-images.githubusercontent.com/12729184/114728510-bc3e5900-9d71-11eb-9721-8a414dca21a0.png)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }

                                    var cellValue = prop.GetValue(item);
                                    if (cellValue == null)
                                    {
                                        rowXml.Replace(key, "");
                                        continue;
                                    }

                                    var cellValueStr = ExcelOpenXmlUtils.EncodeXML(cellValue?.ToString());
                                    var type = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                                    if (type == typeof(bool))
                                    {
                                        cellValueStr = (bool)cellValue ? "1" : "0";
                                    }
                                    else if (type == typeof(DateTime))
                                    {
                                        cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                                    }
                                    else if (TypeHelper.IsNumericType(type))
                                    {
                                        if (decimal.TryParse(cellValueStr, out var decimalValue))
                                            cellValueStr = decimalValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    }

                                    //TODO: ![image](https://user-images.githubusercontent.com/12729184/114848248-17735880-9e11-11eb-8258-63266bda0a1a.png)

                                    rowXml.Replace("@header" + key, cellValueStr);
                                    rowXml.Replace(key, cellValueStr);

                                    if (isHeaderRow && row.InnerText.Contains(key))
                                    {
                                        currentHeader += cellValueStr;
                                    }
                                }
                            }

                            if (isHeaderRow)
                            {
                                if (currentHeader == prevHeader)
                                {
                                    headerDiff++;
                                    continue;
                                }
                                else
                                {
                                    prevHeader = currentHeader;
                                }
                            }

                            // note: only first time need add diff ![image](https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png)
                            if (!first)
                                //rowIndexDiff++;
                                rowIndexDiff += (rowInfo.IEnumerableMercell?.Height ?? 1); //TODO:base on the merge size
                            first = false;

                            var mergeBaseRowIndex = newRowIndex;
                            newRowIndex += rowInfo.IEnumerableMercell?.Height ?? 1;

							// replace formulas
							ProcessFormulas( rowXml, newRowIndex );
							writer.Write(CleanXml( rowXml, endPrefix)); // pass StringBuilder for netcoreapp3.0 or above

                            //mergecells
                            if (rowInfo.RowMercells != null)
                            {
                                foreach (var mergeCell in rowInfo.RowMercells)
                                {
                                    var newMergeCell = new XMergeCell(mergeCell);
                                    newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                                    newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                                    this.NewXMergeCellInfos.Add(newMergeCell);
                                }

                                // Last merge one don't add new row, or it'll get duplicate result like : https://github.com/shps951023/MiniExcel/issues/207#issuecomment-824550950
                                if (iEnumerableIndex == rowInfo.CellIEnumerableValuesCount)
                                    continue;

                                if (rowInfo.IEnumerableMercell != null)
                                    continue;

                                // https://github.com/mini-software/MiniExcel/assets/12729184/1a699497-57e8-4602-b01e-9ffcfef1478d
                                if (rowInfo?.IEnumerableMercell?.Height != null)
                                {
                                    // https://github.com/shps951023/MiniExcel/issues/207#issuecomment-824518897
                                    for (int i = 1; i < rowInfo.IEnumerableMercell.Height; i++)
                                    {
                                        mergeBaseRowIndex++;
                                        var _newRow = row.Clone() as XmlElement;
                                        _newRow.SetAttribute("r", mergeBaseRowIndex.ToString());

                                        var cs = _newRow.SelectNodes($"x:c", _ns);
                                        // all v replace by empty
                                        // TODO: remove c/v
                                        foreach (XmlElement _c in cs)
                                        {
                                            _c.RemoveAttribute("t");
                                            foreach (XmlNode ch in _c.ChildNodes)
                                            {
                                                _c.RemoveChild(ch);
                                            }
                                        }

                                        _newRow.InnerXml = new StringBuilder(_newRow.InnerXml).Replace($"{{{{$rowindex}}}}", mergeBaseRowIndex.ToString()).ToString();
                                        writer.Write(CleanXml(_newRow.OuterXml, endPrefix));
                                    }
                                }

                            }
                        }

                        enumrowend = newRowIndex-1;
                    }
                    else
                    {

                        rowXml.Clear()
                              .Append(outerXmlOpen)
                              .AppendFormat(@" r=""{0}"">", newRowIndex)
                              .Append(innerXml)
                              .Replace($"{{{{$rowindex}}}}", newRowIndex.ToString())
                              .Replace($"{{{{$enumrowstart}}}}", enumrowstart.ToString())
                              .Replace($"{{{{$enumrowend}}}}", enumrowend.ToString())
                              .AppendFormat("</{0}>", row.Name);

						ProcessFormulas( rowXml, newRowIndex );

						writer.Write(CleanXml( rowXml, endPrefix)); // pass StringBuilder for netcoreapp3.0 or above

                        //mergecells
                        if (rowInfo.RowMercells != null)
                        {
                            foreach (var mergeCell in rowInfo.RowMercells)
                            {
                                var newMergeCell = new XMergeCell(mergeCell);
                                newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                                newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                                this.NewXMergeCellInfos.Add(newMergeCell);
                            }
                        }

                    }

                    // get the row's all mergecells then update the rowindex
                }
                #endregion

                writer.Write($"</{prefix}sheetData>");

                if (this.NewXMergeCellInfos.Count != 0)
                {
                    writer.Write($"<{prefix}mergeCells count=\"{this.NewXMergeCellInfos.Count}\">");
                    foreach (var cell in this.NewXMergeCellInfos)
                    {
                        writer.Write(cell.ToXmlString(prefix));
                    }
                    writer.Write($"</{prefix}mergeCells>");
                }

                writer.Write(contents[1]);
            }
        }
        
        private void ProcessFormulas( StringBuilder rowXml, int rowIndex )
        {

            var rowXmlString = rowXml.ToString();

            // exit early if possible
            if ( !rowXmlString.Contains( "$=" ) ) {
                return;
            }

			XmlReaderSettings settings = new XmlReaderSettings { NameTable = _ns.NameTable };
			XmlParserContext context = new XmlParserContext( null, _ns, "", XmlSpace.Default );
			XmlReader reader = XmlReader.Create( new StringReader( rowXmlString ), settings, context );
            
            XmlDocument d = new XmlDocument();
            d.Load( reader );

            var row = d.FirstChild as XmlElement;

			// convert cells starting with '$=' into formulas
			var cs = row.SelectNodes( $"x:c", _ns );
			for ( var ci = 0; ci < cs.Count; ci++ )
            {
				var c = cs.Item( ci ) as XmlElement;
				if ( c == null ) {
					continue;
				}
				/* Target:
				 <c r="C8" s="3">
					<f>SUM(C2:C7)</f>
				</c>
				 */
				var vs = c.SelectNodes( $"x:v", _ns );
				foreach ( XmlElement v in vs )
                {
					if ( !v.InnerText.StartsWith( "$=" ) )
                    {
						continue;
					}
					var fNode = c.OwnerDocument.CreateElement( "f", Config.SpreadsheetmlXmlns );
					fNode.InnerText = v.InnerText.Substring( 2 );
					c.InsertBefore( fNode, v );
					c.RemoveChild( v );

					var celRef = ExcelOpenXmlUtils.ConvertXyToCell( ci + 1, rowIndex );
					CalcChainCellRefs.Add( celRef );

				}
			}
            rowXml.Clear();
			rowXml.Append( row.OuterXml );
		}

        private static string ConvertToDateTimeString(KeyValuePair<string, PropInfo> propInfo, object cellValue)
        {
            string cellValueStr;


            //TODO:c.SetAttribute("t", "d"); and custom format
            var format = string.Empty;
            if (propInfo.Value.PropertyInfo == null)
            {
                format = "yyyy-MM-dd HH:mm:ss";
            }
            else
            {
                format = propInfo.Value.PropertyInfo.GetAttributeValue((ExcelFormatAttribute x) => x.Format)
                             ?? propInfo.Value.PropertyInfo.GetAttributeValue((ExcelColumnAttribute x) => x.Format)
                             ?? "yyyy-MM-dd HH:mm:ss";
            }

            cellValueStr = (cellValue as DateTime?)?.ToString(format);
            return cellValueStr;
        }

        private static StringBuilder CleanXml(StringBuilder xml, string endPrefix)
        {
            return xml
               .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
               .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
        }

        private static string CleanXml(string xml, string endPrefix)
        {
            //TODO: need to optimize
            return CleanXml(new StringBuilder(xml), endPrefix)
                .ToString();
        }

        private void ReplaceSharedStringsToStr(IDictionary<int, string> sharedStrings, ref XmlNodeList rows)
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
                        if (sharedStrings.ContainsKey(int.Parse(v.InnerText)))
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

        private void UpdateDimensionAndGetRowsInfo(Dictionary<string, object> inputMaps, ref XmlDocument doc, ref XmlNodeList rows, bool changeRowIndex = true)
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
                var xRowInfo = new XRowInfo
                {
                    Row = row
                };
                XRowInfos.Add(xRowInfo);
                foreach (XmlElement c in row.SelectNodes($"x:c", _ns))
                {
                    var r = c.GetAttribute("r");

                    // ==== mergecells ====
                    if (this.XMergeCellInfos.ContainsKey(r))
                    {
                        if (xRowInfo.RowMercells == null)
                            xRowInfo.RowMercells = new List<XMergeCell>();
                        xRowInfo.RowMercells.Add(this.XMergeCellInfos[r]);
                    }

                    if (changeRowIndex)
                    {
                        c.SetAttribute("r", $"{StringHelper.GetLetter(r)}{{{{$rowindex}}}}"); //TODO:
                    }

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

                        // TODO: default if not contain property key, clean the template string
                        if (!inputMaps.ContainsKey(propNames[0]))
                        {
                            if (_configuration.IgnoreTemplateParameterMissing)
                            {
                                v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", "");
                                break;
                            }
                            else
                            {
                                throw new KeyNotFoundException($"Please check {propNames[0]} parameter, it's not exist.");
                            }
                        }

                        var cellValue = inputMaps[propNames[0]]; // 1. From left to right, only the first set is used as the basis for the list
                        if ((cellValue is IEnumerable || cellValue is IList<object>) && !(cellValue is string))
                        {
                            if (this.XMergeCellInfos.ContainsKey(r))
                            {
                                if (xRowInfo.IEnumerableMercell == null)
                                {
                                    xRowInfo.IEnumerableMercell = this.XMergeCellInfos[r];
                                }
                            }

                            xRowInfo.CellIEnumerableValues = cellValue as IEnumerable;
                            xRowInfo.CellIlListValues = cellValue as IList<object>;

                            // get ienumerable runtime type
                            if (xRowInfo.IEnumerableGenricType == null) //avoid duplicate to add rowindexdiff ![image](https://user-images.githubusercontent.com/12729184/114851348-522ac000-9e14-11eb-8244-4730754d6885.png)
                            {
                                var first = true;
                                //TODO:if CellIEnumerableValues is ICollection or Array then get length or Count

                                foreach (var element in xRowInfo.CellIEnumerableValues) //TODO: optimize performance?
                                {
                                    xRowInfo.CellIEnumerableValuesCount++;

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

                                                var values = new Dictionary<string, PropInfo>();

                                                var props = xRowInfo.IEnumerableGenricType.GetProperties();

                                                foreach (var p in props)
                                                {
                                                    values.Add(p.Name, new PropInfo
                                                    {
                                                        PropertyInfo = p,
                                                        PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.PropertyInfo,
                                                        UnderlyingTypePropType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType
                                                    });
                                                }

                                                var fields = xRowInfo.IEnumerableGenricType.GetFields();
                                                foreach (var f in fields)
                                                {
                                                    if (!values.ContainsKey(f.Name))
                                                    {

                                                        values.Add(f.Name, new PropInfo
                                                        {
                                                            FieldInfo = f,
                                                            PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.FieldInfo,
                                                            UnderlyingTypePropType = Nullable.GetUnderlyingType(f.FieldType) ?? f.FieldType
                                                        });
                                                    }
                                                }

                                                xRowInfo.PropsMap = values;
                                            }
                                        }
                                    // ==== get dimension max rowindex ====
                                    if (!first) //avoid duplicate add first one, this row not add status  ![image](https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                        maxRowIndexDiff = maxRowIndexDiff + (xRowInfo.IEnumerableMercell == null ? 1 : xRowInfo.IEnumerableMercell.Height);
                                    first = false;
                                }
                            }

                            //TODO: check if not contain 1 index
                            //only check first one match IEnumerable, so only render one collection at same row

                            // Empty collection parameter will get exception  https://gitee.com/dotnetchina/MiniExcel/issues/I4WM67
                            if (xRowInfo.PropsMap == null)
                            {
                                v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", propNames[1]);
                                break;
                            }
                            if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                            {
                                v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}.{propNames[1]}}}}}", "");
                                continue;
                                throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");
                            }
                            // auto check type https://github.com/shps951023/MiniExcel/issues/177
                            var prop = xRowInfo.PropsMap[propNames[1]];
                            var type = prop.UnderlyingTypePropType; //avoid nullable

                            if (isMultiMatch)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (TypeHelper.IsNumericType(type))
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
                                xRowInfo.CellIEnumerableValues = dt.Rows.Cast<object>().ToList(); //TODO: need to optimize performance
                                xRowInfo.CellIlListValues = dt.Rows.Cast<object>().ToList();
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
                            else if (TypeHelper.IsNumericType(type))
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
                            var cellValueStr = cellValue?.ToString(); /* value did encodexml, so don't duplicate encode value https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN*/
                            if (isMultiMatch) // if matchs count over 1 need to set type=str ![image](https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                            {
                                c.SetAttribute("t", "str");
                            }
                            else if (decimal.TryParse(cellValueStr, out var outV))
                            {
                                c.SetAttribute("t", "n");
                                cellValueStr = outV.ToString(System.Globalization.CultureInfo.InvariantCulture);
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
            if (@refs.Length == 2)
            {
                var letter = new String(refs[1].Where(Char.IsLetter).ToArray());
                var digit = int.Parse(new String(refs[1].Where(Char.IsDigit).ToArray()));

                dimension.SetAttribute("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
            }
            else
            {
                var letter = new String(refs[0].Where(Char.IsLetter).ToArray());
                var digit = int.Parse(new String(refs[0].Where(Char.IsDigit).ToArray()));

                dimension.SetAttribute("ref", $"A1:{letter}{digit + maxRowIndexDiff}");
            }
        }

        private static bool EvaluateStatement(object tagValue, string comparisonOperator, string value)
        {
            var checkStatement = false;

            switch (tagValue)
            {
                case double dtg when double.TryParse(value, out var doubleNumber):
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = dtg.Equals(doubleNumber);
                            break;
                        case "!=":
                            checkStatement = !dtg.Equals(doubleNumber);
                            break;
                        case ">":
                            checkStatement = dtg > doubleNumber;
                            break;
                        case "<":
                            checkStatement = dtg < doubleNumber;
                            break;
                        case ">=":
                            checkStatement = dtg >= doubleNumber;
                            break;
                        case "<=":
                            checkStatement = dtg <= doubleNumber;
                            break;
                    }

                    break;
                case int itg when int.TryParse(value, out var intNumber):
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = itg.Equals(intNumber);
                            break;
                        case "!=":
                            checkStatement = !itg.Equals(intNumber);
                            break;
                        case ">":
                            checkStatement = itg > intNumber;
                            break;
                        case "<":
                            checkStatement = itg < intNumber;
                            break;
                        case ">=":
                            checkStatement = itg >= intNumber;
                            break;
                        case "<=":
                            checkStatement = itg <= intNumber;
                            break;
                    }

                    break;
                case DateTime dttg when DateTime.TryParse(value, out var date):
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = dttg.Equals(date);
                            break;
                        case "!=":
                            checkStatement = !dttg.Equals(date);
                            break;
                        case ">":
                            checkStatement = dttg > date;
                            break;
                        case "<":
                            checkStatement = dttg < date;
                            break;
                        case ">=":
                            checkStatement = dttg >= date;
                            break;
                        case "<=":
                            checkStatement = dttg <= date;
                            break;
                    }

                    break;
                case string stg:
                    switch (comparisonOperator)
                    {
                        case "==":
                            checkStatement = stg == value;
                            break;
                        case "!=":
                            checkStatement = stg != value;
                            break;
                    }

                    break;
            }

            return checkStatement;
        }
    }
}