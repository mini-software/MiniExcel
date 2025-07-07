using MiniExcelLib.OpenXml.Constants;

namespace MiniExcelLib.OpenXml.Templates;

#region Utils
internal class XRowInfo
{
    public string FormatText { get; set; }
    public string IEnumerablePropName { get; set; }
    public XmlElement Row { get; set; }
    public Type IEnumerableGenericType { get; set; }
    public IDictionary<string, PropInfo> PropsMap { get; set; }
    public bool IsDictionary { get; set; }
    public bool IsDataTable { get; set; }
    public int CellIEnumerableValuesCount { get; set; }
    public IList<object>? CellIlListValues { get; set; }
    public IEnumerable? CellIEnumerableValues { get; set; }
    public XMergeCell? IEnumerableMercell { get; set; }
    public List<XMergeCell>? RowMercells { get; set; }
    public List<XmlElement>? ConditionalFormats { get; set; }


}

internal class PropInfo
{
    public PropertyInfo PropertyInfo { get; set; }
    public FieldInfo FieldInfo { get; set; }
    public Type UnderlyingTypePropType { get; set; }
    public PropertyInfoOrFieldInfo PropertyInfoOrFieldInfo { get; set; } = PropertyInfoOrFieldInfo.None;
}

internal enum PropertyInfoOrFieldInfo
{
    None = 0,
    PropertyInfo = 1,
    FieldInfo = 2
}

internal class XMergeCell
{
    public XMergeCell(XMergeCell mergeCell)
    {
        Width = mergeCell.Width;
        Height = mergeCell.Height;
        X1 = mergeCell.X1;
        Y1 = mergeCell.Y1;
        X2 = mergeCell.X2;
        Y2 = mergeCell.Y2;
        MergeCell = mergeCell.MergeCell;
    }
    public XMergeCell(XmlElement mergeCell)
    {
        var refAttr = mergeCell.Attributes["ref"].Value;
        var refs = refAttr.Split(':');

        //TODO: width,height
        var xy1 = refs[0];
        X1 = ColumnHelper.GetColumnIndex(StringHelper.GetLetters(refs[0]));
        Y1 = StringHelper.GetNumber(xy1);

        var xy2 = refs[1];
        X2 = ColumnHelper.GetColumnIndex(StringHelper.GetLetters(refs[1]));
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

    public string XY1 => $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}";
    public int X1 { get; set; }
    public int Y1 { get; set; }
    public string XY2 => $"{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}";
    public int X2 { get; set; }
    public int Y2 { get; set; }
    public string Ref => $"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}";
    public XmlElement MergeCell { get; set; }
    public int Width { get; internal set; }
    public int Height { get; internal set; }

    public string ToXmlString(string prefix)
        => $"<{prefix}mergeCell ref=\"{ColumnHelper.GetAlphabetColumnName(X1)}{Y1}:{ColumnHelper.GetAlphabetColumnName(X2)}{Y2}\"/>";
}

internal class MergeCellIndex(int rowStart, int rowEnd)
{
    public int RowStart { get; } = rowStart;
    public int RowEnd { get; } = rowEnd;
}

internal class XChildNode
{
    public string? InnerText { get; set; }
    public string ColIndex { get; set; }
    public int RowIndex { get; set; }
}

internal struct Range
{
    public int StartColumn { get; set; }
    public int StartRow { get; set; }
    public int EndColumn { get; set; }
    public int EndRow { get; set; }

    public bool ContainsRow(int row) => StartRow <= row && row <= EndRow;
}

internal class ConditionalFormatRange
{
    public XmlNode? Node { get; set; }
    public List<Range> Ranges { get; set; } = [];
}
#endregion

internal partial class OpenXmlTemplate
{
    private List<XRowInfo> _xRowInfos;
    private readonly List<string> _calcChainCellRefs = [];
    private Dictionary<string, XMergeCell> _xMergeCellInfos;
    private List<XMergeCell> _newXMergeCellInfos;

#if NET7_0_OR_GREATER
    [GeneratedRegex("([A-Z]+)([0-9]+)")] private static partial Regex CellRegexImpl();
    private static readonly Regex CellRegex = CellRegexImpl();
    [GeneratedRegex(@"\{\{(.*?)\}\}")] private static partial Regex TemplateRegexImpl();
    private static readonly Regex TemplateRegex = TemplateRegexImpl();
    [GeneratedRegex(@".*?\{\{.*?\}\}.*?")] private static partial Regex NonTemplateRegexImpl();
    private static readonly Regex NonTemplateRegex = NonTemplateRegexImpl();
#else
    private static readonly Regex CellRegex = new("([A-Z]+)([0-9]+)", RegexOptions.Compiled);
    private static readonly Regex TemplateRegex = new(@"\{\{(.*?)\}\}", RegexOptions.Compiled);
    private static readonly Regex NonTemplateRegex = new(@".*?\{\{.*?\}\}.*?", RegexOptions.Compiled);
#endif

    [CreateSyncVersion]
    private async Task GenerateSheetXmlImplByUpdateModeAsync(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream, IDictionary<string, object> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false, CancellationToken cancellationToken = default)
    {
        var doc = new XmlDocument();
        doc.Load(sheetStream);
#if NET5_0_OR_GREATER
            await sheetStream.DisposeAsync().ConfigureAwait(false);
#else
        sheetStream.Dispose();
#endif

        sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

        var worksheet = doc.SelectSingleNode("/x:worksheet", Ns);
        var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", Ns);
        var newSheetData = sheetData?.Clone(); //avoid delete lost data
        var rows = newSheetData?.SelectNodes("x:row", Ns);

        ReplaceSharedStringsToStr(sharedStrings, rows);
        GetMergeCells(doc, worksheet);
        UpdateDimensionAndGetRowsInfo(inputMaps, doc, rows, !mergeCells);

        await WriteSheetXmlAsync(stream, doc, sheetData, mergeCells, cancellationToken).ConfigureAwait(false);
    }
    
    private void GenerateSheetXmlImplByCreateMode(ZipArchiveEntry templateSheetZipEntry, Stream outputZipSheetEntryStream, Stream outputSheetStream, IDictionary<string, object> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false)
    {
        var doc = new XmlDocument();
        using (var newTemplateStream = templateSheetZipEntry.Open())
        {
            doc.Load(newTemplateStream);
        }

        //outputSheetStream.Dispose();
        //sheetZipEntry.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic

        var worksheet = doc.SelectSingleNode("/x:worksheet", Ns);
        var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", Ns);
        var newSheetData = sheetData?.Clone(); //avoid delete lost data
        var rows = newSheetData?.SelectNodes("x:row", Ns);

        ReplaceSharedStringsToStr(sharedStrings, rows);
        GetMergeCells(doc, worksheet);
        UpdateDimensionAndGetRowsInfo(inputMaps, doc, rows, !mergeCells);

        WriteSheetXml(outputZipSheetEntryStream, doc, sheetData, mergeCells);
    }

    private void GetMergeCells(XmlDocument doc, XmlNode worksheet)
    {
        var mergeCells = doc.SelectSingleNode("/x:worksheet/x:mergeCells", Ns);
        if (mergeCells is null)
            return;

        var newMergeCells = mergeCells.Clone();
        worksheet.RemoveChild(mergeCells);

        foreach (XmlElement cell in newMergeCells)
        {
            var mergerCell = new XMergeCell(cell);
            _xMergeCellInfos[mergerCell.XY1] = mergerCell;
        }
    }

    private static IEnumerable<ConditionalFormatRange> ParseConditionalFormatRanges(XmlDocument doc)
    {
        var conditionalFormatting = doc.SelectNodes("/x:worksheet/x:conditionalFormatting", Ns);
        if (conditionalFormatting is null)
            yield break;

        foreach (XmlNode conditionalFormat in conditionalFormatting)
        {
            var rangeValues = conditionalFormat.Attributes?["sqref"]?.Value.Split(' ');
            if (rangeValues is null)
                continue;

            var rangeList = new List<Range>();
            foreach (var rangeVal in rangeValues)
            {
                var rangeValSplit = rangeVal.Split(':');
                if (rangeValSplit.Length == 0)
                    continue;

                if (rangeValSplit.Length == 1)
                {
                    var match = CellRegex.Match(rangeValSplit[0]);
                    if (!match.Success)
                        continue;

                    var row = int.Parse(match.Groups[2].Value);
                    var column = ColumnHelper.GetColumnIndex(match.Groups[1].Value);
                    rangeList.Add(new Range
                    {
                        StartColumn = column,
                        StartRow = row,
                        EndColumn = column,
                        EndRow = row
                    });
                }
                else
                {
                    var match1 = CellRegex.Match(rangeValSplit[0]);
                    var match2 = CellRegex.Match(rangeValSplit[1]);
                    if (match1.Success && match2.Success)
                    {
                        rangeList.Add(new Range
                        {
                            StartColumn = ColumnHelper.GetColumnIndex(match1.Groups[1].Value),
                            StartRow = int.Parse(match1.Groups[2].Value),
                            EndColumn = ColumnHelper.GetColumnIndex(match2.Groups[1].Value),
                            EndRow = int.Parse(match2.Groups[2].Value)
                        });
                    }
                }
            }

            yield return new ConditionalFormatRange
            {
                Node = conditionalFormat,
                Ranges = rangeList
            };
        }
    }

    [CreateSyncVersion]
    private async Task WriteSheetXmlAsync(Stream outputFileStream, XmlDocument doc, XmlNode sheetData, bool mergeCells = false, CancellationToken cancellationToken = default)
    {
        //Q.Why so complex?
        //A.Because try to use string stream avoid OOM when rendering rows

        var conditionalFormatRanges = ParseConditionalFormatRanges(doc).ToList();
        var newConditionalFormatRanges = new List<ConditionalFormatRange>();
        newConditionalFormatRanges.AddRange(conditionalFormatRanges);

        sheetData.RemoveAll();
        sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad code smell

        var prefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $"{sheetData.Prefix}:";
        var endPrefix = string.IsNullOrEmpty(sheetData.Prefix) ? "" : $":{sheetData.Prefix}"; // https://user-images.githubusercontent.com/12729184/115000066-fd02b300-9ed4-11eb-8e65-bf0014015134.png

        var conditionalFormatNodes = doc.SelectNodes("/x:worksheet/x:conditionalFormatting", Ns);
        for (var i = 0; i < conditionalFormatNodes?.Count; ++i)
        {
            var node = conditionalFormatNodes.Item(i);
            node.ParentNode.RemoveChild(node);
        }

        var phoneticPr = doc.SelectSingleNode("/x:worksheet/x:phoneticPr", Ns);
        var phoneticPrXml = string.Empty;
        if (phoneticPr is not null)
        {
            phoneticPrXml = phoneticPr.OuterXml;
            phoneticPr.ParentNode.RemoveChild(phoneticPr);
        }

        var contents = doc.InnerXml.Split(new[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None);

        using var writer = new StreamWriter(outputFileStream, Encoding.UTF8);
        await writer.WriteAsync(contents[0]
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
        ).ConfigureAwait(false);
        await writer.WriteAsync($"<{prefix}sheetData>"
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
        ).ConfigureAwait(false); // prefix problem

        if (mergeCells)
        {
            MergeCells(_xRowInfos);
        }

        #region Generate rows and cells

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
        string prevHeader = "";
        int mergeRowCount = 0;

        for (int rowNo = 0; rowNo < _xRowInfos.Count; rowNo++)
        {
            var isHeaderRow = false;
            var currentHeader = "";

            var rowInfo = _xRowInfos[rowNo];
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
            else if (mergeCells)
            {
                if (row.InnerText.Contains("@merge") || row.InnerText.Contains("@endmerge"))
                {
                    mergeRowCount++;
                    continue;
                }
            }

            if (groupingStarted && !isCellIEnumerableValuesSet)
            {
                cellIEnumerableValues = rowInfo.CellIlListValues
                                        ?? rowInfo.CellIEnumerableValues.Cast<object>().ToList();
                isCellIEnumerableValuesSet = true;
            }

            var groupingRowDiff = hasEverGroupStarted
                ? cellIEnumerableValuesIndex * groupRowCount - headerDiff - 1 : 0;

            if (groupingStarted)
            {
                if (isFirstRound)
                {
                    groupRowCount++;
                }

                if (cellIEnumerableValues is not null)
                {
                    rowInfo.CellIEnumerableValuesCount = 1;

                    var listValue = new List<object> { cellIEnumerableValues[cellIEnumerableValuesIndex] };
                    rowInfo.CellIEnumerableValues = listValue;
                    rowInfo.CellIlListValues = listValue;
                }
            }

            //TODO: some xlsx without r
            var originRowIndex = int.Parse(row.GetAttribute("r"));
            var newRowIndex = originRowIndex + rowIndexDiff + groupingRowDiff - mergeRowCount;

            string innerXml = row.InnerXml;
            rowXml.Clear().AppendFormat("<{0}", row.Name);
            foreach (XmlAttribute attr in row.Attributes)
            {
                if (attr.Name != "r")
                    rowXml.AppendFormat(@" {0}=""{1}""", attr.Name, attr.Value);
            }

            var outerXmlOpen = new StringBuilder();
            outerXmlOpen.Append(rowXml);

            if (rowInfo.CellIEnumerableValues is not null)
            {
                var isFirst = true;
                var iEnumerableIndex = 0;
                enumrowstart = newRowIndex;

                var generateCellValuesContext = new GenerateCellValuesContext()
                {
                    currentHeader = currentHeader,
                    headerDiff = headerDiff,
                    iEnumerableIndex = iEnumerableIndex,
                    isFirst = isFirst,
                    newRowIndex = newRowIndex,
                    prevHeader = prevHeader,
                    rowIndexDiff = rowIndexDiff,
                };

                generateCellValuesContext = await GenerateCellValuesAsync(generateCellValuesContext, endPrefix, writer, rowXml, mergeRowCount, isHeaderRow, rowInfo, row, groupingRowDiff, innerXml, outerXmlOpen, row, cancellationToken).ConfigureAwait(false);

                rowIndexDiff = generateCellValuesContext.rowIndexDiff;
                headerDiff = generateCellValuesContext.headerDiff;
                prevHeader = generateCellValuesContext.prevHeader;
                newRowIndex = generateCellValuesContext.newRowIndex;
                isFirst = generateCellValuesContext.isFirst;
                iEnumerableIndex = generateCellValuesContext.iEnumerableIndex;
                currentHeader = generateCellValuesContext.currentHeader;

                enumrowend = newRowIndex - 1;

                var conditionalFormats = conditionalFormatRanges.Where(cfr => cfr.Ranges.Any(r => r.ContainsRow(originRowIndex)));
                foreach (var conditionalFormat in conditionalFormats)
                {
                    var newConditionalFormat = conditionalFormat.Node.Clone();
                    var sqref = newConditionalFormat.Attributes["sqref"];
                    var ranges = conditionalFormat.Ranges
                        .Where(r => r.ContainsRow(originRowIndex))
                        .Select(r => new Range
                        {
                            StartColumn = r.StartColumn,
                            StartRow = enumrowstart,
                            EndColumn = r.EndColumn,
                            EndRow = enumrowend
                        })
                        .ToList();

                    sqref.Value = string.Join(" ", ranges.Select(r => $"{ColumnHelper.GetAlphabetColumnName(r.StartColumn)}{r.StartRow}:{ColumnHelper.GetAlphabetColumnName(r.EndColumn)}{r.EndRow}"));
                    newConditionalFormatRanges.Remove(conditionalFormat);
                    newConditionalFormatRanges.Add(new ConditionalFormatRange
                    {
                        Node = newConditionalFormat,
                        Ranges = ranges
                    });
                }
            }
            else
            {
                rowXml.Clear()
                    .Append(outerXmlOpen)
                    .AppendFormat(@" r=""{0}"">", newRowIndex)
                    .Append(innerXml)
                    .Replace("{{$rowindex}}", newRowIndex.ToString())
                    .Replace("{{$enumrowstart}}", enumrowstart.ToString())
                    .Replace("{{$enumrowend}}", enumrowend.ToString())
                    .AppendFormat("</{0}>", row.Name);

                ProcessFormulas(rowXml, newRowIndex);
                await writer.WriteAsync(CleanXml(rowXml, endPrefix).ToString()
#if NET5_0_OR_GREATER
                            .AsMemory(), cancellationToken
#endif
                ).ConfigureAwait(false);

                //mergecells
                if (rowInfo.RowMercells is null)
                    continue;

                foreach (var mergeCell in rowInfo.RowMercells)
                {
                    var newMergeCell = new XMergeCell(mergeCell);
                    newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                    newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                    _newXMergeCellInfos.Add(newMergeCell);
                }
            }
            // get the row's all mergecells then update the rowindex
        }

        #endregion

        await writer.WriteAsync($"</{prefix}sheetData>"
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
        ).ConfigureAwait(false);

        if (_newXMergeCellInfos.Count != 0)
        {
            await writer.WriteAsync($"<{prefix}mergeCells count=\"{_newXMergeCellInfos.Count}\">"
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
            foreach (var cell in _newXMergeCellInfos)
            {
                await writer.WriteAsync(cell.ToXmlString(prefix)
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                ).ConfigureAwait(false);
            }
            await writer.WriteLineAsync($"</{prefix}mergeCells>"
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        if (!string.IsNullOrEmpty(phoneticPrXml))
        {
            await writer.WriteAsync(phoneticPrXml
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        if (newConditionalFormatRanges.Count != 0)
        {
            await writer.WriteAsync(string.Join(string.Empty, newConditionalFormatRanges.Select(cf => cf.Node.OuterXml))
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        await writer.WriteAsync(contents[1]
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
        ).ConfigureAwait(false);
    }

    class GenerateCellValuesContext
    {
        public int rowIndexDiff { get; set; }
        public int headerDiff { get; set; }
        public string prevHeader { get; set; }
        public string currentHeader { get; set; }
        public int newRowIndex { get; set; }
        public bool isFirst { get; set; }
        public int iEnumerableIndex { get; set; }
    }

    //todo: refactor in a way that needs less parameters
    [CreateSyncVersion]
    private async Task<GenerateCellValuesContext> GenerateCellValuesAsync(GenerateCellValuesContext generateCellValuesContext, string endPrefix, StreamWriter writer,
        StringBuilder rowXml, int mergeRowCount, bool isHeaderRow,
        XRowInfo rowInfo, XmlElement row, int groupingRowDiff,
        string innerXml, StringBuilder outerXmlOpen, XmlElement rowElement, CancellationToken cancellationToken = default)
    {
        var rowIndexDiff = generateCellValuesContext.rowIndexDiff;
        var headerDiff = generateCellValuesContext.headerDiff;
        var prevHeader = generateCellValuesContext.prevHeader;
        var newRowIndex = generateCellValuesContext.newRowIndex;
        var isFirst = generateCellValuesContext.isFirst;
        var iEnumerableIndex = generateCellValuesContext.iEnumerableIndex;
        var currentHeader = generateCellValuesContext.currentHeader;

        // Just need to remove space string one time https://github.com/mini-software/MiniExcel/issues/751
        var cleanOuterXmlOpen = CleanXml(outerXmlOpen, endPrefix);
        var cleanInnerXml = CleanXml(innerXml, endPrefix);

        // https://github.com/mini-software/MiniExcel/issues/771 Saving by template introduces unintended value replication in each row #771
        var notFirstRowElement = rowElement.Clone();
        foreach (XmlElement c in notFirstRowElement.SelectNodes("x:c", Ns))
        {
            var v = c.SelectSingleNode("x:v", Ns);
            if (v is not null && !NonTemplateRegex.IsMatch(v.InnerText))
                v.InnerText = string.Empty;
        }
        var cleanNotFirstRowInnerXml = CleanXml(notFirstRowElement.InnerXml, endPrefix);

        foreach (var item in rowInfo.CellIEnumerableValues)
        {
            iEnumerableIndex++;
            rowXml.Clear()
                .Append(cleanOuterXmlOpen)
                .AppendFormat(@" r=""{0}"">", newRowIndex)
                .Append(cleanInnerXml)
                .Replace("{{$rowindex}}", newRowIndex.ToString())
                .AppendFormat("</{0}>", row.Name);

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

            var isDictOrTable = rowInfo.IsDictionary || rowInfo.IsDataTable;
            var dict = item as IDictionary<string, object>;
            var dataRow = item as DataRow;

            for (var i = 0; i < lines.Length; i++)
            {
                if (lines[i].Contains("@if") || lines[i].Contains("@elseif"))
                {
                    var newLines = lines[i]
                        .Replace("@elseif(", "")
                        .Replace("@if(", "")
                        .Replace(")", "")
                        .Split(' ');

                    object value;
                    if (rowInfo.IsDictionary)
                    {
                        value = dict[newLines[0]];
                    }
                    else if (rowInfo.IsDataTable)
                    {
                        value = dataRow[newLines[0]];
                    }
                    else
                    {
                        value = string.Empty;
                        var prop = rowInfo.PropsMap[newLines[0]];
                        if (prop.PropertyInfoOrFieldInfo == PropertyInfoOrFieldInfo.PropertyInfo)
                        {
                            value = prop.PropertyInfo.GetValue(item);
                        }
                        else if (prop.PropertyInfoOrFieldInfo == PropertyInfoOrFieldInfo.FieldInfo)
                        {
                            value = prop.FieldInfo.GetValue(item);
                        }
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

            var substXmlRow = rowXml.ToString();

            if (item is null)
            {
                substXmlRow = TemplateRegex.Replace(substXmlRow, "");
            }
            else
            {
                var replacements = new Dictionary<string, string>();
#if NETCOREAPP3_0_OR_GREATER
                    string MatchDelegate(Match x) => CollectionExtensions.GetValueOrDefault(replacements, x.Groups[1].Value, "");
#else
                string MatchDelegate(Match x) => replacements.TryGetValue(x.Groups[1].Value, out var repl) ? repl : "";
#endif
                foreach (var prop in rowInfo.PropsMap)
                {
                    var propInfo = prop.Value.PropertyInfo;
                    var name = isDictOrTable ? prop.Key : propInfo.Name;
                    var key = $"{rowInfo.IEnumerablePropName}.{name}";

                    object? cellValue;
                    if (rowInfo.IsDictionary)
                    {
                        if (!dict.TryGetValue(prop.Key, out cellValue))
                            continue;
                    }
                    else if (rowInfo.IsDataTable)
                    {
                        cellValue = dataRow[prop.Key];
                    }
                    else
                    {
                        cellValue = propInfo.GetValue(item);
                    }

                    if (cellValue is null)
                        continue;

                    var type = isDictOrTable
                        ? prop.Value.UnderlyingTypePropType
                        : Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                    string cellValueStr;
                    if (type == typeof(bool))
                    {
                        cellValueStr = (bool)cellValue ? "1" : "0";
                    }
                    else if (type == typeof(DateTime))
                    {
                        cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                    }
                    else if (type?.IsEnum ?? false)
                    {
                        var description = CustomPropertyHelper.DescriptionAttr(type, cellValue);
                        cellValueStr = XmlHelper.EncodeXml(description);
                    }
                    else
                    {
                        cellValueStr = XmlHelper.EncodeXml(cellValue?.ToString());
                        if (!isDictOrTable && TypeHelper.IsNumericType(type))
                        {
                            if (decimal.TryParse(cellValueStr, out var decimalValue))
                                cellValueStr = decimalValue.ToString(CultureInfo.InvariantCulture);
                        }
                    }

                    replacements[key] = cellValueStr;
                    rowXml.Replace($"@header{{{{{key}}}}}", cellValueStr);

                    if (isHeaderRow && row.InnerText.Contains(key))
                    {
                        currentHeader += cellValueStr;
                    }
                }

                substXmlRow = rowXml.ToString();
                substXmlRow = TemplateRegex.Replace(substXmlRow, MatchDelegate);
            }

            rowXml.Clear();
            rowXml.Append(substXmlRow);

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

            // note: only first time need add diff https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png
            if (!isFirst)
                rowIndexDiff += rowInfo.IEnumerableMercell?.Height ?? 1; //TODO:base on the merge size

            if (isFirst)
            {
                // https://github.com/mini-software/MiniExcel/issues/771 Saving by template introduces unintended value replication in each row #771
                cleanInnerXml = cleanNotFirstRowInnerXml;
                isFirst = false;
            }

            var mergeBaseRowIndex = newRowIndex;
            newRowIndex += rowInfo.IEnumerableMercell?.Height ?? 1;

            // replace formulas
            ProcessFormulas(rowXml, newRowIndex);
            await writer.WriteAsync(rowXml.ToString()
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);

            //mergecells
            if (rowInfo.RowMercells is null)
                continue;

            foreach (var mergeCell in rowInfo.RowMercells)
            {
                var newMergeCell = new XMergeCell(mergeCell);
                newMergeCell.Y1 = newMergeCell.Y1 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                newMergeCell.Y2 = newMergeCell.Y2 + rowIndexDiff + groupingRowDiff - mergeRowCount;
                _newXMergeCellInfos.Add(newMergeCell);
            }

            // Last merge one don't add new row, or it'll get duplicate result like : https://github.com/mini-software/MiniExcel/issues/207#issuecomment-824550950
            if (iEnumerableIndex == rowInfo.CellIEnumerableValuesCount)
                continue;

            if (rowInfo.IEnumerableMercell is not null)
                continue;

            // https://github.com/mini-software/MiniExcel/assets/12729184/1a699497-57e8-4602-b01e-9ffcfef1478d
            if (rowInfo.IEnumerableMercell?.Height is null)
                continue;

            // https://github.com/mini-software/MiniExcel/issues/207#issuecomment-824518897
            for (int i = 1; i < rowInfo.IEnumerableMercell.Height; i++)
            {
                mergeBaseRowIndex++;
                var newRow = row.Clone() as XmlElement;
                newRow.SetAttribute("r", mergeBaseRowIndex.ToString());

                var cs = newRow.SelectNodes("x:c", Ns);
                // all v replace by empty
                // TODO: remove c/v
                foreach (XmlElement c in cs)
                {
                    c.RemoveAttribute("t");
                    foreach (XmlNode ch in c.ChildNodes)
                    {
                        c.RemoveChild(ch);
                    }
                }

                newRow.InnerXml = new StringBuilder(newRow.InnerXml).Replace("{{$rowindex}}", mergeBaseRowIndex.ToString()).ToString();
                await writer.WriteAsync(CleanXml(newRow.OuterXml, endPrefix)
#if NET7_0_OR_GREATER
                    .AsMemory(), cancellationToken
#endif
                ).ConfigureAwait(false);
            }
        }

        return new GenerateCellValuesContext()
        {
            currentHeader = currentHeader,
            headerDiff = headerDiff,
            iEnumerableIndex = iEnumerableIndex,
            isFirst = isFirst,
            newRowIndex = newRowIndex,
            prevHeader = prevHeader,
            rowIndexDiff = rowIndexDiff,
        };
    }

    private static void MergeCells(List<XRowInfo> xRowInfos)
    {
        var mergeTaggedColumns = new Dictionary<XChildNode, XChildNode>();
        var columns = xRowInfos
            .SelectMany(s => s.Row.Cast<XmlElement>())
            .Where(s => !string.IsNullOrEmpty(s.InnerText))
            .Select(s =>
            {
                var att = s.GetAttribute("r");
                return new XChildNode
                {
                    InnerText = s.InnerText,
                    ColIndex = StringHelper.GetLetters(att),
                    RowIndex = StringHelper.GetNumber(att)
                };
            })
            .OrderBy(x => x.RowIndex)
            .ToList();

        var mergeColumns = columns.Where(s => s.InnerText.Contains("@merge")).ToList();
        var endMergeColumns = columns.Where(s => s.InnerText.Contains("@endmerge")).ToList();
        var mergeLimitColumn = mergeColumns.FirstOrDefault(x => x.InnerText.Contains("@mergelimit"));

        foreach (var mergeColumn in mergeColumns)
        {
            var endMergeColumn = endMergeColumns.FirstOrDefault(s =>
                s.ColIndex == mergeColumn.ColIndex && s.RowIndex > mergeColumn.RowIndex);

            if (endMergeColumn is not null)
            {
                mergeTaggedColumns[mergeColumn] = endMergeColumn;
            }
        }

        if (mergeTaggedColumns.Count <= 0)
            return;

        var calculatedColumns = new List<XChildNode>();
        foreach (var taggedColumn in mergeTaggedColumns)
        {
            calculatedColumns.AddRange(columns.Where(x =>
                x.ColIndex == taggedColumn.Key.ColIndex && x.RowIndex > taggedColumn.Key.RowIndex &&
                x.RowIndex < taggedColumn.Value.RowIndex));
        }

        var lastMergeCellIndexes = new Dictionary<int, MergeCellIndex>();
        foreach (var rowInfo in xRowInfos)
        {
            var row = rowInfo.Row;
            var childNodes = row.ChildNodes.Cast<XmlElement>();

            foreach (var childNode in childNodes)
            {
                var att = childNode.GetAttribute("r");
                var childNodeLetter = StringHelper.GetLetters(att);
                var childNodeNumber = StringHelper.GetNumber(att);

                if (!string.IsNullOrEmpty(childNode.InnerText))
                {
                    var xmlNodes = calculatedColumns
                        .Where(j =>
                            j.InnerText == childNode.InnerText &&
                            j.ColIndex == childNodeLetter)
                        .OrderBy(s => s.RowIndex)
                        .ToList();

                    if (xmlNodes.Count > 1)
                    {
                        if (mergeLimitColumn is not null)
                        {
                            var limitedNode = calculatedColumns.First(j =>
                                j.ColIndex == mergeLimitColumn.ColIndex && j.RowIndex == childNodeNumber);

                            var limitedMaxNode = calculatedColumns.Last(j =>
                                j.ColIndex == mergeLimitColumn.ColIndex && j.InnerText == limitedNode.InnerText);

                            xmlNodes = xmlNodes
                                .Where(j =>
                                    j.RowIndex >= limitedNode.RowIndex &&
                                    j.RowIndex <= limitedMaxNode.RowIndex)
                                .ToList();
                        }

                        var firstRow = xmlNodes.FirstOrDefault();
                        var lastRow = xmlNodes.LastOrDefault(s =>
                            s.RowIndex <= firstRow?.RowIndex + xmlNodes.Count &&
                            s.RowIndex != firstRow?.RowIndex);

                        if (firstRow is not null && lastRow is not null)
                        {
                            var mergeCell = new XMergeCell(firstRow.ColIndex, firstRow.RowIndex, lastRow.ColIndex, lastRow.RowIndex);
                            var mergeIndexResult = lastMergeCellIndexes.TryGetValue(mergeCell.X1, out var mergeIndex);

                            if (!mergeIndexResult ||
                                mergeCell.Y1 < mergeIndex?.RowStart ||
                                mergeCell.Y2 > mergeIndex?.RowEnd)
                            {
                                lastMergeCellIndexes[mergeCell.X1] = new MergeCellIndex(mergeCell.Y1, mergeCell.Y2);

                                rowInfo.RowMercells ??= [];
                                rowInfo.RowMercells.Add(mergeCell);
                            }
                        }
                    }
                }

                childNode.SetAttribute("r", $"{childNodeLetter}{{{{$rowindex}}}}");
            }
        }
    }

    private void ProcessFormulas(StringBuilder rowXml, int rowIndex)
    {
        var rowXmlString = rowXml.ToString();

        // exit early if possible
        if (!rowXmlString.Contains("$="))
            return;

        var settings = new XmlReaderSettings { NameTable = Ns.NameTable };
        var context = new XmlParserContext(null, Ns, "", XmlSpace.Default);
        
        using var reader = XmlReader.Create(new StringReader(rowXmlString), settings, context);
        var d = new XmlDocument();
        d.Load(reader);

        var row = d.FirstChild as XmlElement;

        // convert cells starting with '$=' into formulas
        var cs = row.SelectNodes("x:c", Ns);
        for (var ci = 0; ci < cs.Count; ci++)
        {
            var c = cs.Item(ci) as XmlElement;
            if (c is null)
                continue;

            /* Target:
                 <c r="C8" s="3">
                    <f>SUM(C2:C7)</f>
                </c>
                 */
            var vs = c.SelectNodes("x:v", Ns);
            foreach (XmlElement v in vs)
            {
                if (!v.InnerText.StartsWith("$="))
                    continue;

                var fNode = c.OwnerDocument.CreateElement("f", Schemas.SpreadsheetmlXmlns);
                fNode.InnerText = v.InnerText[2..];
                c.InsertBefore(fNode, v);
                c.RemoveChild(v);

                var celRef = ReferenceHelper.ConvertCoordinatesToCell(ci + 1, rowIndex);
                _calcChainCellRefs.Add(celRef);
            }
        }

        rowXml.Clear();
        rowXml.Append(row.OuterXml);
    }

    private static string? ConvertToDateTimeString(PropertyInfo? propInfo, object cellValue)
    {
        //TODO:c.SetAttribute("t", "d"); and custom format
        var format = propInfo?.GetAttributeValue((MiniExcelFormatAttribute x) => x.Format)
                     ?? propInfo?.GetAttributeValue((MiniExcelColumnAttribute x) => x.Format)
                     ?? "yyyy-MM-dd HH:mm:ss";

        return (cellValue as DateTime?)?.ToString(format);
    }

    //TODO: need to optimize
    private static string CleanXml(string xml, string endPrefix) => CleanXml(new StringBuilder(xml), endPrefix).ToString();
    private static StringBuilder CleanXml(StringBuilder xml, string endPrefix) => xml
        .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
        .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");

    private static void ReplaceSharedStringsToStr(IDictionary<int, string> sharedStrings, XmlNodeList rows)
    {
        foreach (XmlElement row in rows)
        {
            var cs = row.SelectNodes("x:c", Ns);
            foreach (XmlElement c in cs)
            {
                var t = c.GetAttribute("t");
                var v = c.SelectSingleNode("x:v", Ns);
                if (v?.InnerText is null) //![image](https://user-images.githubusercontent.com/12729184/114363496-075a3f80-9bab-11eb-9883-8e3fec10765c.png)
                    continue;

                if (t != "s")
                    continue;

                //need to check sharedstring exist or not
                if (sharedStrings is null || !sharedStrings.TryGetValue(int.Parse(v.InnerText), out var shared))
                    continue;

                // change type = str and replace its value
                //TODO: remove sharedstring?
                v.InnerText = shared;
                c.SetAttribute("t", "str");
            }
        }
    }

    private void UpdateDimensionAndGetRowsInfo(IDictionary<string, object> inputMaps, XmlDocument doc, XmlNodeList rows, bool changeRowIndex = true)
    {
        string[] refs;
        if (doc.SelectSingleNode("/x:worksheet/x:dimension", Ns) is XmlElement dimension)
        {
            refs = dimension.GetAttribute("ref").Split(':');
        }
        else
        {
            // ==== add dimension element if not found ====

            var firstRow = rows[0].SelectNodes("x:c", Ns);
            var lastRow = rows[^1].SelectNodes("x:c", Ns);

            var dimStart = ((XmlElement)firstRow?[0])?.GetAttribute("r");
            var dimEnd = ((XmlElement)lastRow?[lastRow.Count - 1])?.GetAttribute("r");

            refs = new[] { dimStart, dimEnd };

            dimension = (XmlElement)doc.CreateNode(XmlNodeType.Element, "dimension", null);
            var worksheet = doc.SelectSingleNode("/x:worksheet", Ns);
            worksheet?.InsertBefore(dimension, worksheet.FirstChild);
        }

        var maxRowIndexDiff = 0;
        foreach (XmlElement row in rows)
        {
            // ==== get ienumerable infomation & maxrowindexdiff ====

            var xRowInfo = new XRowInfo { Row = row };
            _xRowInfos.Add(xRowInfo);

            foreach (XmlElement c in row.SelectNodes("x:c", Ns))
            {
                var r = c.GetAttribute("r");

                // ==== mergecells ====
                if (_xMergeCellInfos.TryGetValue(r, out var merCell))
                {
                    if (xRowInfo.RowMercells is null)
                        xRowInfo.RowMercells = new List<XMergeCell>();
                    xRowInfo.RowMercells.Add(merCell);
                }

                if (changeRowIndex)
                {
                    c.SetAttribute("r", $"{StringHelper.GetLetters(r)}{{{{$rowindex}}}}");
                }

                var v = c.SelectSingleNode("x:v", Ns);
                if (v?.InnerText is null)
                    continue;

                var matches = IsExpressionRegex.Matches(v.InnerText)
                    .Cast<Match>()
                    .Select(x => x.Value)
                    .Distinct()
                    .ToArray();

                var matchCount = matches.Length;
                var isMultiMatch = matchCount > 1 || (matchCount == 1 && v.InnerText != $"{{{{{matches[0]}}}}}");

                foreach (var formatText in matches)
                {
                    xRowInfo.FormatText = formatText;
                    var propNames = formatText.Split('.');
                    if (propNames[0].StartsWith("$")) //e.g:"$rowindex" it doesn't need to check cell value type
                        continue;

                    // TODO: default if not contain property key, clean the template string
                    if (!inputMaps.TryGetValue(propNames[0], out var cellValue))
                    {
                        if (!_configuration.IgnoreTemplateParameterMissing)
                            throw new KeyNotFoundException($"The parameter '{propNames[0]}' was not found.");

                        v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", "");
                        break;
                    }

                    //cellValue = inputMaps[propNames[0]] - 1. From left to right, only the first set is used as the basis for the list
                    if (cellValue is IEnumerable value && !(cellValue is string))
                    {
                        if (xRowInfo.IEnumerableMercell is null && _xMergeCellInfos.TryGetValue(r, out var info))
                        {
                            xRowInfo.IEnumerableMercell = info;
                        }

                        xRowInfo.CellIEnumerableValues = value;

                        // get ienumerable runtime type
                        if (xRowInfo.IEnumerableGenericType is null) //avoid duplicate to add rowindexdiff (https://user-images.githubusercontent.com/12729184/114851348-522ac000-9e14-11eb-8244-4730754d6885.png)
                        {
                            //TODO: optimize performance?              

                            var first = true;
                            foreach (var element in xRowInfo.CellIEnumerableValues)
                            {
                                xRowInfo.CellIEnumerableValuesCount++;
                                if (xRowInfo.IEnumerableGenericType is null && element is not null)
                                {
                                    xRowInfo.IEnumerablePropName = propNames[0];
                                    xRowInfo.IEnumerableGenericType = element.GetType();

                                    if (element is IDictionary<string, object> dic)
                                    {
                                        xRowInfo.IsDictionary = true;
                                        xRowInfo.PropsMap = dic.ToDictionary(
                                            kv => kv.Key,
                                            kv => kv.Value is not null
                                                ? new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(kv.Value.GetType()) ?? kv.Value.GetType() }
                                                : new PropInfo { UnderlyingTypePropType = typeof(object) });
                                    }
                                    else
                                    {
                                        var props = xRowInfo.IEnumerableGenericType.GetProperties();
                                        var values = props.ToDictionary(
                                            p => p.Name,
                                            p => new PropInfo
                                            {
                                                PropertyInfo = p,
                                                PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.PropertyInfo,
                                                UnderlyingTypePropType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType
                                            });

                                        var fields = xRowInfo.IEnumerableGenericType.GetFields();
                                        foreach (var f in fields)
                                        {
                                            if (!values.ContainsKey(f.Name))
                                            {
                                                var propInfo = new PropInfo
                                                {
                                                    FieldInfo = f,
                                                    PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.FieldInfo,
                                                    UnderlyingTypePropType = Nullable.GetUnderlyingType(f.FieldType) ?? f.FieldType
                                                };
                                                values.Add(f.Name, propInfo);
                                            }
                                        }

                                        xRowInfo.PropsMap = values;
                                    }
                                }

                                // ==== get dimension max rowindex ====
                                if (!first) //avoid duplicate add first one, this row not add status  (https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                    maxRowIndexDiff += xRowInfo.IEnumerableMercell?.Height ?? 1;
                                first = false;
                            }
                        }

                        //TODO: check if not contain 1 index
                        //only check first one match IEnumerable, so only render one collection at same row

                        // Empty collection parameter will get exception  https://gitee.com/dotnetchina/MiniExcel/issues/I4WM67
                        if (xRowInfo.PropsMap is null)
                        {
                            v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", propNames[1]);
                            break;
                        }
                        if (!xRowInfo.PropsMap.TryGetValue(propNames[1], out var prop))
                        {
                            v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}.{propNames[1]}}}}}", "");
                            continue;

                            //why unreachable exception?
                            throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");
                        }
                        // auto check type https://github.com/mini-software/MiniExcel/issues/177
                        var type = prop.UnderlyingTypePropType; //avoid nullable

                        if (isMultiMatch)
                        {
                            c.SetAttribute("t", "str");
                        }
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
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
                    else if (cellValue is DataTable dt)
                    {
                        if (xRowInfo.CellIEnumerableValues is null)
                        {
                            xRowInfo.IEnumerablePropName = propNames[0];
                            xRowInfo.IEnumerableGenericType = typeof(DataRow);
                            xRowInfo.IsDataTable = true;

                            var listValues = dt.Rows.Cast<object>().ToList();
                            xRowInfo.CellIEnumerableValues = listValues;
                            xRowInfo.CellIlListValues = listValues;

                            var first = true;
                            foreach (var element in xRowInfo.CellIEnumerableValues)
                            {
                                // ==== get demension max rowindex ====
                                if (!first) //avoid duplicate add first one, this row not add status (https://user-images.githubusercontent.com/12729184/114851829-d2512580-9e14-11eb-8e7d-520c89a7ebee.png)
                                    maxRowIndexDiff++;
                                first = false;
                            }
                            //TODO:need to optimize
                            //maxRowIndexDiff = dt.Rows.Count <= 1 ? 0 : dt.Rows.Count-1;
                            xRowInfo.PropsMap = dt.Columns.Cast<DataColumn>().ToDictionary(col => 
                                col.ColumnName,
                                col => new PropInfo { UnderlyingTypePropType = Nullable.GetUnderlyingType(col.DataType) }
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
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
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
                        var cellValueStr = cellValue?.ToString(); // value did encodexml, so don't duplicate encode value (https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)
                        if (isMultiMatch || cellValue is string) // if matchs count over 1 need to set type=str (https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                        {
                            c.SetAttribute("t", "str");
                        }
                        else if (decimal.TryParse(cellValueStr, out var outV))
                        {
                            c.SetAttribute("t", "n");
                            cellValueStr = outV.ToString(CultureInfo.InvariantCulture);
                        }
                        else if (cellValue is bool b)
                        {
                            c.SetAttribute("t", "b");
                            cellValueStr = b ? "1" : "0";
                        }
                        else if (cellValue is DateTime timestamp)
                        {
                            //c.SetAttribute("t", "d");
                            cellValueStr = timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                        }

                        v.InnerText = v.InnerText.Replace($"{{{{{propNames[0]}}}}}", cellValueStr); //TODO: auto check type and set value
                    }
                }
                //if (xRowInfo.CellIEnumerableValues is not null) //2. From left to right, only the first set is used as the basis for the list
                //    break;
            }
        }

        // e.g <dimension ref=\"A1:B6\" /> we only need to update B6 to BMaxRowIndex
        if (refs.Length == 2)
        {
            var letter = StringHelper.GetLetters(refs[1]);
            var digit = StringHelper.GetNumber(refs[1]);
            dimension.SetAttribute("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
        }
        else
        {
            var letter = StringHelper.GetLetters(refs[0]);
            var digit = StringHelper.GetNumber(refs[0]);
            dimension.SetAttribute("ref", $"A1:{letter}{digit + maxRowIndexDiff}");
        }
    }

    private static bool EvaluateStatement(object tagValue, string comparisonOperator, string value)
    {
        return tagValue switch
        {
            double dtg when double.TryParse(value, out var doubleNumber) => comparisonOperator switch
            {
                "==" => dtg.Equals(doubleNumber),
                "!=" => !dtg.Equals(doubleNumber),
                ">" => dtg > doubleNumber,
                "<" => dtg < doubleNumber,
                ">=" => dtg >= doubleNumber,
                "<=" => dtg <= doubleNumber
            },
            
            int itg when int.TryParse(value, out var intNumber) => comparisonOperator switch
            {
                "==" => itg.Equals(intNumber),
                "!=" => !itg.Equals(intNumber),
                ">" => itg > intNumber,
                "<" => itg < intNumber,
                ">=" => itg >= intNumber,
                "<=" => itg <= intNumber
            },
            
            DateTime dttg when DateTime.TryParse(value, out var date) => comparisonOperator switch
            {
                "==" => dttg.Equals(date),
                "!=" => !dttg.Equals(date),
                ">" => dttg > date,
                "<" => dttg < date,
                ">=" => dttg >= date,
                "<=" => dttg <= date
            },
            
            string stg => comparisonOperator switch
            {
                "==" => stg == value,
                "!=" => stg != value
            },
            
            _ => false
        };
    }
}