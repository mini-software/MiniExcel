using MiniExcelLib.Core.Attributes;
using MiniExcelLib.OpenXml.Constants;
using System.ComponentModel;

namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    private readonly List<string> _calcChainCellRefs = [];
    
    private List<XRowInfo> _xRowInfos;
    private Dictionary<string, XMergeCell> _xMergeCellInfos;
    private List<XMergeCell> _newXMergeCellInfos;

#if NET7_0_OR_GREATER
    [GeneratedRegex("([A-Z]+)([0-9]+)")] private static partial Regex CellRegexImpl();
    private static readonly Regex CellRegex = CellRegexImpl();
    [GeneratedRegex(@"\{\{(.*?)\}\}")] private static partial Regex TemplateRegexImpl();
    private static readonly Regex TemplateRegex = TemplateRegexImpl();
    [GeneratedRegex(@".*?\{\{.*?\}\}.*?")] private static partial Regex NonTemplateRegexImpl();
    private static readonly Regex NonTemplateRegex = NonTemplateRegexImpl();
    [GeneratedRegex(@"<(?:x:)?v>\s*</(?:x:)?v>")] private static partial Regex EmptyVTagRegexImpl();
    private static readonly Regex EmptyVTagRegex = EmptyVTagRegexImpl();
#else
    private static readonly Regex CellRegex = new("([A-Z]+)([0-9]+)", RegexOptions.Compiled);
    private static readonly Regex TemplateRegex = new(@"\{\{(.*?)\}\}", RegexOptions.Compiled);
    private static readonly Regex NonTemplateRegex = new(@".*?\{\{.*?\}\}.*?", RegexOptions.Compiled);
    private static readonly Regex EmptyVTagRegex = new(@"<(?:x:)?v>\s*</(?:x:)?v>", RegexOptions.Compiled);
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
    
    [CreateSyncVersion]
    private async Task GenerateSheetXmlImplByCreateModeAsync(ZipArchiveEntry templateSheetZipEntry, Stream outputZipSheetEntryStream, IDictionary<string, object?> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false)
    {
        var doc = new XmlDocument
        {
            XmlResolver = null
        };
        
#if NET5_0_OR_GREATER
#if NET10_0_OR_GREATER
        var newTemplateStream = await templateSheetZipEntry.OpenAsync().ConfigureAwait(false);
#else
        var newTemplateStream = templateSheetZipEntry.Open();
#endif
        await using var disposableStream = newTemplateStream.ConfigureAwait(false);
#else
        using var newTemplateStream = templateSheetZipEntry.Open();
#endif
        doc.Load(newTemplateStream);

        var worksheet = doc.SelectSingleNode("/x:worksheet", Ns);
        var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", Ns);
        var newSheetData = sheetData?.Clone(); //avoid delete lost data
        var rows = newSheetData?.SelectNodes("x:row", Ns);

        ReplaceSharedStringsToStr(sharedStrings, rows);
        GetMergeCells(doc, worksheet);
        UpdateDimensionAndGetRowsInfo(inputMaps, doc, rows, !mergeCells);

        await WriteSheetXmlAsync(outputZipSheetEntryStream, doc, sheetData, mergeCells).ConfigureAwait(false);
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
                    var column = CellReferenceConverter.GetNumericalIndex(match.Groups[1].Value);
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
                            StartColumn = CellReferenceConverter.GetNumericalIndex(match1.Groups[1].Value),
                            StartRow = int.Parse(match1.Groups[2].Value),
                            EndColumn = CellReferenceConverter.GetNumericalIndex(match2.Groups[1].Value),
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

        // Extract autoFilter - must be written before mergeCells and phoneticPr per ECMA-376
        var autoFilter = doc.SelectSingleNode("/x:worksheet/x:autoFilter", Ns);
        var autoFilterXml = string.Empty;
        if (autoFilter is not null)
        {
            autoFilterXml = autoFilter.OuterXml;
            autoFilter.ParentNode.RemoveChild(autoFilter);
        }

        var contents = doc.InnerXml.Split(new[] { $"<{prefix}sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</{prefix}sheetData>" }, StringSplitOptions.None);
#if NETCOREAPP3_0_OR_GREATER
        var writer = new StreamWriter(outputFileStream, Encoding.UTF8);
        await using var disposableWriter =  writer.ConfigureAwait(false);
#else
        using var writer = new StreamWriter(outputFileStream, Encoding.UTF8);
#endif
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

            SpecialCellType specialCellType = default;
            foreach (XmlNode c in row.GetElementsByTagName("c"))
            {
                specialCellType = c.InnerText switch
                {
                    "@group" => SpecialCellType.Group,
                    "@endgroup" => SpecialCellType.Endgroup,
                    "@merge" or "@endmerge" => SpecialCellType.Merge,
                    var s when s.StartsWith("@header") => SpecialCellType.Header,
                    _ => SpecialCellType.None
                };

                if (specialCellType != SpecialCellType.None)
                    break;
            }
            
            if (specialCellType == SpecialCellType.Group)
            {
                groupingStarted = true;
                hasEverGroupStarted = true;
                groupStartRowIndex = rowNo;
                isFirstRound = true;
                prevHeader = "";
                continue;
            }
            else if (specialCellType == SpecialCellType.Endgroup)
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
            else if (specialCellType == SpecialCellType.Header)
            {
                isHeaderRow = true;
            }
            else if (mergeCells)
            {
                if (specialCellType == SpecialCellType.Merge)
                {
                    mergeRowCount++;
                    continue;
                }
            }

            if (groupingStarted && !isCellIEnumerableValuesSet)
            {
                cellIEnumerableValues = rowInfo.CellIlListValues ?? rowInfo.CellIEnumerableValues.Cast<object>().ToList();
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
                    CurrentHeader = currentHeader,
                    HeaderDiff = headerDiff,
                    EnumerableIndex = iEnumerableIndex,
                    IsFirst = isFirst,
                    NewRowIndex = newRowIndex,
                    PrevHeader = prevHeader,
                    RowIndexDiff = rowIndexDiff,
                };

                generateCellValuesContext = await GenerateCellValuesAsync(generateCellValuesContext, endPrefix, writer, rowXml, mergeRowCount, isHeaderRow, rowInfo, row, groupingRowDiff, innerXml, outerXmlOpen, row, cancellationToken).ConfigureAwait(false);

                rowIndexDiff = generateCellValuesContext.RowIndexDiff;
                headerDiff = generateCellValuesContext.HeaderDiff;
                prevHeader = generateCellValuesContext.PrevHeader;
                newRowIndex = generateCellValuesContext.NewRowIndex;
                isFirst = generateCellValuesContext.IsFirst;
                iEnumerableIndex = generateCellValuesContext.EnumerableIndex;
                currentHeader = generateCellValuesContext.CurrentHeader;

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

                    sqref.Value = string.Join(" ", ranges.Select(r => $"{CellReferenceConverter.GetAlphabeticalIndex(r.StartColumn)}{r.StartRow}:{CellReferenceConverter.GetAlphabeticalIndex(r.EndColumn)}{r.EndRow}"));
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

        // ECMA-376 element order: sheetData → autoFilter → mergeCells → phoneticPr → conditionalFormatting
        
        // 1. autoFilter (must come before mergeCells)
        if (!string.IsNullOrEmpty(autoFilterXml))
        {
            await writer.WriteAsync(CleanXml(autoFilterXml, endPrefix)
#if NET7_0_OR_GREATER
                .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        // 2. mergeCells
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

        // 3. phoneticPr
        if (!string.IsNullOrEmpty(phoneticPrXml))
        {
            await writer.WriteAsync(CleanXml(phoneticPrXml, endPrefix)
#if NET7_0_OR_GREATER
                .AsMemory(), cancellationToken
#endif
            ).ConfigureAwait(false);
        }

        // 4. conditionalFormatting
        if (newConditionalFormatRanges.Count != 0)
        {
            await writer.WriteAsync(CleanXml(string.Join(string.Empty, newConditionalFormatRanges.Select(cf => cf.Node.OuterXml)), endPrefix)
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

    //todo: refactor in a way that needs less parameters
    [CreateSyncVersion]
    private async Task<GenerateCellValuesContext> GenerateCellValuesAsync(
        GenerateCellValuesContext generateCellValuesContext, 
        string endPrefix, 
        StreamWriter writer,
        StringBuilder rowXml, 
        int mergeRowCount, 
        bool isHeaderRow, 
        XRowInfo rowInfo, 
        XmlElement row, 
        int groupingRowDiff,
        string innerXml, 
        StringBuilder outerXmlOpen, 
        XmlElement rowElement, 
        CancellationToken cancellationToken = default)
    {
        var rowIndexDiff = generateCellValuesContext.RowIndexDiff;
        var headerDiff = generateCellValuesContext.HeaderDiff;
        var prevHeader = generateCellValuesContext.PrevHeader;
        var newRowIndex = generateCellValuesContext.NewRowIndex;
        var isFirst = generateCellValuesContext.IsFirst;
        var iEnumerableIndex = generateCellValuesContext.EnumerableIndex;
        var currentHeader = generateCellValuesContext.CurrentHeader;
        
        // https://github.com/mini-software/MiniExcel/issues/771 Saving by template introduces unintended value replication in each row #771
        var notFirstRowElement = rowElement.Clone();
        foreach (XmlElement c in notFirstRowElement.SelectNodes("x:c", Ns))
        {
            // Try <v> first (for t="n"/t="b" cells), then <is><t> (for t="inlineStr" cells)
            var vTag = c.SelectSingleNode("x:v", Ns);
            if (vTag is not null)
            {
                if (!NonTemplateRegex.IsMatch(vTag.InnerText))
                    vTag.InnerText = string.Empty;
            }
            else
            {
                // Handle inline string cells
                var t = c.SelectSingleNode("x:is/x:t", Ns);
                if (t is not null && !NonTemplateRegex.IsMatch(t.InnerText))
                    t.InnerText = string.Empty;
            }
        }

        foreach (var item in rowInfo.CellIEnumerableValues)
        {
            iEnumerableIndex++;
            rowXml.Clear()
                .Append(outerXmlOpen)
                .AppendFormat(@" r=""{0}"">", newRowIndex)
                .Append(innerXml)
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

                    object? value;
                    if (rowInfo.IsDictionary)
                    {
                        value = dict![newLines[0]];
                    }
                    else if (rowInfo.IsDataTable)
                    {
                        value = dataRow![newLines[0]];
                    }
                    else
                    {
                        var map = rowInfo.MembersMap[newLines[0]];
                        value = map.PropertyInfoOrFieldInfo switch
                        {
                            PropertyInfoOrFieldInfo.PropertyInfo => map.PropertyInfo.GetValue(item),
                            PropertyInfoOrFieldInfo.FieldInfo => map.FieldInfo.GetValue(item),
                            _ => string.Empty
                        };
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
                string MatchDelegate(Match x) => replacements.GetValueOrDefault(x.Groups[1].Value, "");

                foreach (var map in rowInfo.MembersMap)
                {
                    var propInfo = map.Value.PropertyInfo;
                    var name = isDictOrTable ? map.Key : propInfo.Name;
                    var key = $"{rowInfo.IEnumerablePropName}.{name}";

                    object? cellValue;
                    if (rowInfo.IsDictionary)
                    {
                        if (!dict!.TryGetValue(map.Key, out cellValue))
                            continue;
                    }
                    else if (rowInfo.IsDataTable)
                    {
                        cellValue = dataRow![map.Key];
                    }
                    else
                    {
                        cellValue = propInfo.GetValue(item);
                    }

                    if (cellValue is null)
                        continue;

                    var type = isDictOrTable
                        ? map.Value.UnderlyingMemberType
                        : Nullable.GetUnderlyingType(propInfo.PropertyType) ?? propInfo.PropertyType;

                    string? cellValueStr;
                    if (type == typeof(bool))
                    {
                        cellValueStr = (bool)cellValue ? "1" : "0";
                    }
                    else if (type == typeof(DateTime))
                    {
                        cellValueStr = ConvertToDateTimeString(propInfo, cellValue);
                    }
                    else if (type?.IsEnum is true)
                    {
                        var stringValue = Enum.GetName(type, cellValue) ?? "";

                        var attr = type.GetField(stringValue)?.GetCustomAttribute<DescriptionAttribute>();
                        var description = attr?.Description ?? stringValue;

                        cellValueStr = XmlHelper.EncodeXml(description);
                    }
                    else
                    {
                        cellValueStr = XmlHelper.EncodeXml(cellValue?.ToString());
                        if (TypeHelper.IsNumericType(type))
                        {
                            if (decimal.TryParse(cellValueStr, out var decimalValue))
                                cellValueStr = decimalValue.ToString(CultureInfo.InvariantCulture);
                        }
                    }

                    // escaping formulas
                    var tempReplacement = cellValueStr ?? "";
                    var replacementValue = tempReplacement.StartsWith("$=") || tempReplacement.StartsWith("=")
                        ? $"&apos;{tempReplacement}" 
                        : tempReplacement;
                    
                    replacements[key] = replacementValue;
                    rowXml.Replace($"@header{{{{{key}}}}}", replacementValue);

                    if (isHeaderRow && row.InnerText.Contains(key))
                    {
                        currentHeader += cellValueStr;
                    }
                }

                substXmlRow = rowXml.ToString();
                substXmlRow = TemplateRegex.Replace(substXmlRow, MatchDelegate);
                
                // Cleanup empty <v> tags which defaults to invalid XML
                substXmlRow = EmptyVTagRegex.Replace(substXmlRow, "");
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
                innerXml = notFirstRowElement.InnerXml;
                isFirst = false;
            }

            var mergeBaseRowIndex = newRowIndex;
            newRowIndex += rowInfo.IEnumerableMercell?.Height ?? 1;

            // Replace {{$rowindex}} in the already-built substXmlRow
            rowXml.Replace("{{$rowindex}}", mergeBaseRowIndex.ToString());

            // replace formulas
            ProcessFormulas(rowXml, newRowIndex);
            var finalXml = CleanXml(rowXml, endPrefix).ToString();
            await writer.WriteAsync(finalXml
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

            // https://github.com/mini-software/MiniExcel/assets/12729184/1a699497-57e8-4602-b01e-9ffcfef1478d
            if (rowInfo.IEnumerableMercell?.Height is not { } height)
                continue;

            // https://github.com/mini-software/MiniExcel/issues/207#issuecomment-824518897
            for (int i = 1; i < height; i++)
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

        return new GenerateCellValuesContext
        {
            CurrentHeader = currentHeader,
            HeaderDiff = headerDiff,
            EnumerableIndex = iEnumerableIndex,
            IsFirst = isFirst,
            NewRowIndex = newRowIndex,
            PrevHeader = prevHeader,
            RowIndexDiff = rowIndexDiff,
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
            var vs = c.SelectNodes("x:is", Ns);
            foreach (XmlElement v in vs)
            {
                if (!v.InnerText.StartsWith("$="))
                    continue;

                var fNode = c.OwnerDocument.CreateElement("f", Schemas.SpreadsheetmlXmlns);
                fNode.InnerText = v.InnerText[2..];
                c.InsertBefore(fNode, v);
                c.RemoveChild(v);

                var celRef = CellReferenceConverter.GetCellFromCoordinates(ci + 1, rowIndex);
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
        .Replace($"xmlns{endPrefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "")
        .Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");

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

                // change type = inlineStr and replace its value
                // Use the same prefix as the source element to handle namespaced documents (e.g., x:v -> x:is, x:t)
                var prefix = v.Prefix;
                c.RemoveChild(v);
                var isNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("is", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "is", Schemas.SpreadsheetmlXmlns);
                var tNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("t", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "t", Schemas.SpreadsheetmlXmlns);
                tNode.InnerText = shared;
                isNode.AppendChild(tNode);
                c.AppendChild(isNode);

                c.RemoveAttribute("t");
                c.SetAttribute("t", "inlineStr");
            }
        }
    }

    private static void SetCellType(XmlElement c, string type)
    {
        if (type == "str") type = "inlineStr"; // Force inlineStr for strings

        // Determine the prefix used in this document (e.g., "x" for x:c, x:v, etc.)
        var prefix = c.Prefix;

        if (type == "inlineStr")
        {
            // Ensure <is><t>...</t></is>
            c.SetAttribute("t", "inlineStr");
            var v = c.SelectSingleNode("x:v", Ns);
            if (v != null)
            {
                var text = v.InnerText;
                c.RemoveChild(v);
                var isNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("is", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "is", Schemas.SpreadsheetmlXmlns);
                var tNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("t", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "t", Schemas.SpreadsheetmlXmlns);
                tNode.InnerText = text;
                isNode.AppendChild(tNode);
                c.AppendChild(isNode);
            }
            else if (c.SelectSingleNode("x:is", Ns) == null)
            {
                // Create empty <is><t></t></is> if neither <v> nor <is> exists
                var isNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("is", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "is", Schemas.SpreadsheetmlXmlns);
                var tNode = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("t", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "t", Schemas.SpreadsheetmlXmlns);
                isNode.AppendChild(tNode);
                c.AppendChild(isNode);
            }
        }
        else
        {
            // Ensure <v>...</v>
            // For numbers/booleans, we remove 't' attribute to let it be default (number) 
            // or we could set it to 'n' explicitly, but removing is safer for general number types
            if (type == "b")
                c.SetAttribute("t", "b");
            else
                c.RemoveAttribute("t"); 

            var isNode = c.SelectSingleNode("x:is", Ns);
            if (isNode != null)
            {
                var tNode = isNode.SelectSingleNode("x:t", Ns);
                var text = tNode?.InnerText ?? string.Empty;
                c.RemoveChild(isNode);
                var v = string.IsNullOrEmpty(prefix)
                    ? c.OwnerDocument.CreateElement("v", Schemas.SpreadsheetmlXmlns)
                    : c.OwnerDocument.CreateElement(prefix, "v", Schemas.SpreadsheetmlXmlns);
                v.InnerText = text;
                c.AppendChild(v);
            }
        }
    }

    private void UpdateDimensionAndGetRowsInfo(IDictionary<string, object?> inputMaps, XmlDocument doc, XmlNodeList rows, bool changeRowIndex = true)
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

            var dimStart = ((XmlElement?)firstRow?[0])?.GetAttribute("r") ?? "";
            var dimEnd = ((XmlElement?)lastRow?[^1])?.GetAttribute("r") ?? "";

            refs = [dimStart, dimEnd];

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
                    xRowInfo.RowMercells ??= [];
                    xRowInfo.RowMercells.Add(merCell);
                }

                if (changeRowIndex)
                {
                    c.SetAttribute("r", $"{StringHelper.GetLetters(r)}{{{{$rowindex}}}}");
                }

                var v = c.SelectSingleNode("x:v", Ns) ?? c.SelectSingleNode("x:is/x:t", Ns);
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
                    var mapNames = formatText.Split('.');
                    if (mapNames[0].StartsWith("$")) //e.g:"$rowindex" it doesn't need to check cell value type
                        continue;

                    // TODO: default if not contain property key, clean the template string
                    if (!inputMaps.TryGetValue(mapNames[0], out var cellValue))
                    {
                        if (!_configuration.IgnoreTemplateParameterMissing)
                            throw new KeyNotFoundException($"The parameter '{mapNames[0]}' was not found.");

                        v.InnerText = v.InnerText.Replace($"{{{{{mapNames[0]}}}}}", "");
                        break;
                    }

                    //cellValue = inputMaps[propNames[0]] - 1. From left to right, only the first set is used as the basis for the list
                    if (cellValue is IEnumerable value and not string)
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
                                    xRowInfo.IEnumerablePropName = mapNames[0];
                                    xRowInfo.IEnumerableGenericType = element.GetType();

                                    if (element is IDictionary<string, object> dic)
                                    {
                                        xRowInfo.IsDictionary = true;
                                        xRowInfo.MembersMap = dic.ToDictionary(
                                            kv => kv.Key,
                                            kv => kv.Value is not null
                                                ? new MemberInfo { UnderlyingMemberType = Nullable.GetUnderlyingType(kv.Value.GetType()) ?? kv.Value.GetType() }
                                                : new MemberInfo { UnderlyingMemberType = typeof(object) });
                                    }
                                    else
                                    {
                                        var props = xRowInfo.IEnumerableGenericType.GetProperties();
                                        var values = props.ToDictionary(
                                            p => p.Name,
                                            p => new MemberInfo
                                            {
                                                PropertyInfo = p,
                                                PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.PropertyInfo,
                                                UnderlyingMemberType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType
                                            });

                                        var fields = xRowInfo.IEnumerableGenericType.GetFields();
                                        foreach (var f in fields)
                                        {
                                            if (!values.ContainsKey(f.Name))
                                            {
                                                var fieldInfo = new MemberInfo
                                                {
                                                    FieldInfo = f,
                                                    PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.FieldInfo,
                                                    UnderlyingMemberType = Nullable.GetUnderlyingType(f.FieldType) ?? f.FieldType
                                                };
                                                values.Add(f.Name, fieldInfo);
                                            }
                                        }

                                        xRowInfo.MembersMap = values;
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
                        if (xRowInfo.MembersMap is null)
                        {
                            v.InnerText = v.InnerText.Replace($"{{{{{mapNames[0]}}}}}", mapNames[1]);
                            break;
                        }
                        if (!xRowInfo.MembersMap.TryGetValue(mapNames[1], out var map))
                        {
                            v.InnerText = v.InnerText.Replace($"{{{{{mapNames[0]}.{mapNames[1]}}}}}", "");
                            continue;

                            //why unreachable exception?
                            throw new InvalidDataException($"{mapNames[0]} doesn't have {mapNames[1]} property");
                        }
                        // auto check type https://github.com/mini-software/MiniExcel/issues/177
                        var type = map.UnderlyingMemberType; //avoid nullable

                        if (isMultiMatch)
                        {
                            SetCellType(c, "str");
                        }
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
                        {
                            SetCellType(c, "n");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                        {
                            SetCellType(c, "b");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                        {
                            SetCellType(c, "str");
                        }

                        break;
                    }
                    else if (cellValue is DataTable dt)
                    {
                        if (xRowInfo.CellIEnumerableValues is null)
                        {
                            xRowInfo.IEnumerablePropName = mapNames[0];
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
                            xRowInfo.MembersMap = dt.Columns.Cast<DataColumn>().ToDictionary(col => 
                                col.ColumnName,
                                col => new MemberInfo { UnderlyingMemberType = Nullable.GetUnderlyingType(col.DataType) }
                            );
                        }

                        var column = dt.Columns[mapNames[1]];
                        var type = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType; //avoid nullable
                        if (!xRowInfo.MembersMap.ContainsKey(mapNames[1]))
                            throw new InvalidDataException($"{mapNames[0]} doesn't have {mapNames[1]} property");

                        if (isMultiMatch)
                        {
                            SetCellType(c, "str");
                        }
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
                        {
                            SetCellType(c, "n");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                        {
                            SetCellType(c, "b");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                        {
                            SetCellType(c, "str");
                        }
                    }
                    else
                    {
                        var cellValueStr = cellValue?.ToString(); // value did encodexml, so don't duplicate encode value (https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)
                        if (isMultiMatch || cellValue is string) // if matchs count over 1 need to set type=str (https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                        {
                            SetCellType(c, "str");
                        }
                        else if (decimal.TryParse(cellValueStr, out var outV))
                        {
                            SetCellType(c, "n");
                            cellValueStr = outV.ToString(CultureInfo.InvariantCulture);
                        }
                        else if (cellValue is bool b)
                        {
                            SetCellType(c, "b");
                            cellValueStr = b ? "1" : "0";
                        }
                        else if (cellValue is DateTime timestamp)
                        {
                            //c.SetAttribute("t", "d");
                            cellValueStr = timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                        }

                        if (string.IsNullOrEmpty(cellValueStr) && string.IsNullOrEmpty(c.GetAttribute("t")))
                        {
                            SetCellType(c, "str");
                        }

                        // Re-acquire v after SetCellType may have changed DOM structure
                        v = c.SelectSingleNode("x:v", Ns) ?? c.SelectSingleNode("x:is/x:t", Ns);
                        v.InnerText = v.InnerText.Replace($"{{{{{mapNames[0]}}}}}", cellValueStr); //TODO: auto check type and set value
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
