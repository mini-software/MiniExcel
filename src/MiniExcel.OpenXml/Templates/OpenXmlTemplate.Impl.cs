using MiniExcelLib.Core.Attributes;
using System.ComponentModel;
using System.Xml.Linq;
using  MiniExcelLib.OpenXml.Constants;

namespace MiniExcelLib.OpenXml.Templates;

internal partial class OpenXmlTemplate
{
    private static readonly XNamespace SpreadsheetNs = Schemas.SpreadsheetmlXmlMain;

    private static readonly XmlWriterSettings DocXmlWriterSettings = new()
    {
#if !SYNC_ONLY
        Async = true
#endif
    };
    
    private static readonly XmlWriterSettings FragXmlWriterSettings = new()
    {
        OmitXmlDeclaration = true,
        ConformanceLevel =  ConformanceLevel.Fragment,
#if !SYNC_ONLY
        Async = true
#endif
    };

#if NET8_0_OR_GREATER
    [GeneratedRegex("(?<={{).*?(?=}})")] private static partial Regex ExpressionRegex();
    private static readonly Regex IsExpressionRegex = ExpressionRegex();
    [GeneratedRegex("([A-Z]+)([0-9]+)")] private static partial Regex CellRegexImpl();
    private static readonly Regex CellRegex = CellRegexImpl();
    [GeneratedRegex(@"\{\{(.*?)\}\}")] private static partial Regex TemplateRegexImpl();
    private static readonly Regex TemplateRegex = TemplateRegexImpl();
    [GeneratedRegex(@".*?\{\{.*?\}\}.*?")] private static partial Regex NonTemplateRegexImpl();
    private static readonly Regex NonTemplateRegex = NonTemplateRegexImpl();
    [GeneratedRegex(@"<(?:x:)?v>\s*</(?:x:)?v>")] private static partial Regex EmptyVTagRegexImpl();
    private static readonly Regex EmptyVTagRegex = EmptyVTagRegexImpl();
#else
    private static readonly Regex IsExpressionRegex = new("(?<={{).*?(?=}})");
    private static readonly Regex CellRegex = new("([A-Z]+)([0-9]+)", RegexOptions.Compiled);
    private static readonly Regex TemplateRegex = new(@"\{\{(.*?)\}\}", RegexOptions.Compiled);
    private static readonly Regex NonTemplateRegex = new(@".*?\{\{.*?\}\}.*?", RegexOptions.Compiled);
    private static readonly Regex EmptyVTagRegex = new(@"<(?:x:)?v>\s*</(?:x:)?v>", RegexOptions.Compiled);
#endif

    private readonly List<XRowInfo> _xRowInfos = [];
    private readonly Dictionary<string, XMergeCell> _xMergeCellInfos = [];
    private readonly List<XMergeCell> _newXMergeCellInfos = [];
    private readonly List<string> _calcChainCellRefs = [];


    [CreateSyncVersion]
    private async Task GenerateSheetByUpdateModeAsync(ZipArchiveEntry sheetZipEntry, Stream stream, Stream sheetStream, IDictionary<string, object> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var doc = await XDocument.LoadAsync(sheetStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
        await sheetStream.DisposeAsync().ConfigureAwait(false);
#else
        var doc = XDocument.Load(sheetStream);
        sheetStream.Dispose();
#endif

        // we can't update ZipArchiveEntry directly, so we delete the original entry and recreate it
        sheetZipEntry.Delete();

        var worksheet = doc.Element(SpreadsheetNs + "worksheet");
        var sheetData = worksheet?.Element(SpreadsheetNs + "sheetData");
        var newSheetData = new XElement(sheetData);
        var rows = newSheetData.Elements(SpreadsheetNs + "row");

        InjectSharedStrings(sharedStrings, rows);
        GetMergeCells(worksheet);
        UpdateDimensionAndGetRowsInfo(inputMaps, worksheet, rows, !mergeCells);

#if NET8_0_OR_GREATER
        var writer = XmlWriter.Create(stream, DocXmlWriterSettings);
        await using var disposableWriter = writer.ConfigureAwait(false);
#else
        using var writer = XmlWriter.Create(stream, DocXmlWriterSettings);
#endif

        await WriteSheetXmlAsync(writer, worksheet, sheetData, mergeCells, cancellationToken).ConfigureAwait(false);
    }

    [CreateSyncVersion]
    private async Task GenerateSheetByCreateModeAsync(ZipArchiveEntry templateSheetZipEntry, Stream outputZipSheetEntryStream, IDictionary<string, object?> inputMaps, IDictionary<int, string> sharedStrings, bool mergeCells = false, CancellationToken cancellationToken = default)
    {
#if NET8_0_OR_GREATER
        var newTemplateStream = await templateSheetZipEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableNewTemplateStream = newTemplateStream.ConfigureAwait(false);
        var doc = await XDocument.LoadAsync(newTemplateStream, LoadOptions.None, cancellationToken).ConfigureAwait(false);
#else
        using var newTemplateStream = templateSheetZipEntry.Open();
        var doc = XDocument.Load(newTemplateStream);
#endif
        var worksheet = doc.Element(SpreadsheetNs + "worksheet");
        var prefix = worksheet?.GetPrefixOfNamespace(SpreadsheetNs);
        if (!string.IsNullOrEmpty(prefix))
        {
            // we remove the main namespace's prefix declaration so that we don't have to worry about inconstencies when we serialize single elements
            worksheet?.Attribute(XNamespace.Xmlns + prefix)?.Remove();
        }

        var sheetData = worksheet?.Element(SpreadsheetNs + "sheetData");
        var newSheetData = new XElement(sheetData);
    
        var rows = newSheetData.Elements(SpreadsheetNs + "row");

        InjectSharedStrings(sharedStrings, rows);
        GetMergeCells(worksheet);
        UpdateDimensionAndGetRowsInfo(inputMaps, worksheet, rows, !mergeCells);

#if NET8_0_OR_GREATER
        var writer = XmlWriter.Create(outputZipSheetEntryStream, DocXmlWriterSettings);
        await using var disposableWriter = writer.ConfigureAwait(false);
#else
        using var writer = XmlWriter.Create(outputZipSheetEntryStream, DocXmlWriterSettings);
#endif
        await WriteSheetXmlAsync(writer, worksheet, sheetData, mergeCells, cancellationToken).ConfigureAwait(false);
    }

    private void GetMergeCells(XElement worksheet)
    {
        if (worksheet.Element(SpreadsheetNs + "mergeCells") is not { } mergeCells)
            return;

        var newMergeCells = new XElement(mergeCells);
        mergeCells.Remove();

        foreach (var cell in newMergeCells.Elements())
        {
            var mergeCell = new XMergeCell(cell);
            _xMergeCellInfos[mergeCell.XY1] = mergeCell;
        }
    }

    private static IEnumerable<ConditionalFormatRange> NewParseConditionalFormatRanges(XElement worksheet)
    {
        var conditionalFormatting = worksheet.Element(SpreadsheetNs + "conditionalFormatting");
        if (conditionalFormatting is null)
            yield break;

        foreach (var format in conditionalFormatting.Elements())
        {
            var ranges = format.Attribute("sqref")?.Value.Split(' ');
            if (ranges is null)
                continue;

            List<Range> rangeList = [];
            foreach (var range in ranges)
            {
                var rangeValue = range.Split(':');
                if (rangeValue.Length == 0)
                    continue;

                if (rangeValue.Length == 1)
                {
                    if (CellRegex.Match(rangeValue[0]) is not { Success: true } match)
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
                else if (CellRegex.Match(rangeValue[0]) is { Success: true } match1 &&
                         CellRegex.Match(rangeValue[1]) is { Success: true } match2)
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

            yield return new ConditionalFormatRange
            {
                Node = format,
                Ranges = rangeList
            };
        }
    }

    [CreateSyncVersion]
    private async Task WriteSheetXmlAsync(XmlWriter writer, XElement worksheet, XElement sheetData, bool mergeCells = false, CancellationToken cancellationToken = default)
    {
        // TODO: Can we make this less complex?

        var conditionalFormatRanges = NewParseConditionalFormatRanges(worksheet).ToList();
        var newConditionalFormatRanges = new List<ConditionalFormatRange>();
        newConditionalFormatRanges.AddRange(conditionalFormatRanges);

        sheetData.RemoveAll();
        worksheet.Elements(SpreadsheetNs + "conditionalFormatting").Remove();
        
        var prefix = worksheet.GetPrefixOfNamespace(SpreadsheetNs);
        var fullPrefix = !string.IsNullOrEmpty(prefix) ? $"{prefix}:" : "";

        var phoneticPrXml = string.Empty;
        if (worksheet.Element(SpreadsheetNs + "phoneticPr") is { } phoneticPr)
        {
            phoneticPrXml = phoneticPr.ToString(SaveOptions.DisableFormatting);
            phoneticPr.Remove();
        }

        // Extract autoFilter - must be written before mergeCells and phoneticPr per ECMA-376
        var autoFilterXml = string.Empty;
        if (worksheet.Element(SpreadsheetNs + "autoFilter") is { } autoFilter)
        {
            autoFilterXml = autoFilter.ToString(SaveOptions.DisableFormatting);
            autoFilter.Remove();
        }

        var beforeSheetData = worksheet.Element(SpreadsheetNs + "sheetData")?.ElementsBeforeSelf() ?? [];
        var afterSheetData = worksheet.Element(SpreadsheetNs + "sheetData")?.ElementsAfterSelf() ?? [];

        await writer.WriteStartElementAsync(null, "worksheet", Schemas.SpreadsheetmlXmlMain).ConfigureAwait(false);
        foreach (var attr in worksheet.Attributes())
        {
            var (nsPrefix, ns) = attr is { IsNamespaceDeclaration: true, Name.LocalName: not "xmlns" } 
                ? ("xmlns", null as string) 
                : (null, attr.Name.NamespaceName);

            await writer.WriteAttributeStringAsync(nsPrefix, attr.Name.LocalName, ns, attr.Value).ConfigureAwait(false);
        }

        foreach (var beforeElement in beforeSheetData)
        {
#if NET8_0_OR_GREATER
            await beforeElement.WriteToAsync(writer, cancellationToken).ConfigureAwait(false);
#else
            beforeElement.WriteTo(writer);
#endif
        }

        await writer.WriteStartElementAsync(null,"sheetData", Schemas.SpreadsheetmlXmlMain).ConfigureAwait(false);

        if (mergeCells)
        {
            MergeCells(_xRowInfos);
        }

        #region Generate rows and cells

        int rowIndexDiff = 0;

        // for formula cells
        int enumrowstart = -1;
        int enumrowend = -1;

        // for grouped cells
        bool groupingStarted = false;
        bool hasEverGroupStarted = false;
        int groupStartRowIndex = 0;
        IList<object>? cellIEnumerableValues = null;
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
            foreach (var cell in row.Elements(SpreadsheetNs + "c"))
            {
                specialCellType = cell.Value switch
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
                if (cellIEnumerableValuesIndex >= cellIEnumerableValues?.Count - 1)
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
            else if (mergeCells && specialCellType == SpecialCellType.Merge)
            {
                mergeRowCount++;
                continue;
            }

            if (groupingStarted && !isCellIEnumerableValuesSet)
            {
                cellIEnumerableValues = rowInfo.CellIlListValues ?? rowInfo.CellIEnumerableValues?.Cast<object>().ToList() ?? [];
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

            //TODO: Fix parsing for documents that don't have the "r" attribute on rows
            if (row.Attribute("r")?.Value is not { } rVal || !int.TryParse(rVal, out var originRowIndex))
                throw new NotSupportedException("The format of the chosen template is not currently supported.");

            var newRowIndex = originRowIndex + rowIndexDiff + groupingRowDiff - mergeRowCount;
            var innerXml = string.Concat(row.Nodes());

            var rowXml = new StringBuilder("<");
            if (!string.IsNullOrEmpty(prefix)) 
                rowXml.Append($"{prefix}:");
            
            rowXml.Append(row.Name.LocalName);

            foreach (var attr in row.Attributes())
            {
                if (attr is 
                    { 
                        Name: { LocalName: var name and not "r", Namespace: var ns }, 
                        Value: var value 
                    })
                {
                    var pfx = worksheet.GetPrefixOfNamespace(ns);
                    var fullName = string.IsNullOrEmpty(pfx) ? name : $"{pfx}:{name}";
                    rowXml.Append($" {fullName}=\"{value}\"");
                }
            }

            var outerXmlOpen = new StringBuilder().Append(rowXml);
            if (rowInfo.CellIEnumerableValues is not null)
            {
                enumrowstart = newRowIndex;
                var generateCellValuesContext = new GenerateCellValuesContext
                {
                    CurrentHeader = currentHeader,
                    HeaderDiff = headerDiff,
                    EnumerableIndex = 0,
                    IsFirst = true,
                    NewRowIndex = newRowIndex,
                    PrevHeader = prevHeader,
                    RowIndexDiff = rowIndexDiff,
                };

                generateCellValuesContext = await GenerateCellValuesAsync(generateCellValuesContext, prefix, writer, rowXml, mergeRowCount, isHeaderRow, rowInfo, row, groupingRowDiff, innerXml, outerXmlOpen, row, cancellationToken).ConfigureAwait(false);

                rowIndexDiff = generateCellValuesContext.RowIndexDiff;
                headerDiff = generateCellValuesContext.HeaderDiff;
                prevHeader = generateCellValuesContext.PrevHeader;
                newRowIndex = generateCellValuesContext.NewRowIndex;

                enumrowend = newRowIndex - 1;

                var conditionalFormats = conditionalFormatRanges.Where(cfr => cfr.Ranges.Any(r => r.ContainsRow(originRowIndex)));
                foreach (var conditionalFormat in conditionalFormats)
                {
                    var newConditionalFormat = new XElement(conditionalFormat.Node);
                    var sqref = newConditionalFormat.Attribute("sqref");
                    var ranges = conditionalFormat.Ranges
                        .Where(r => r.ContainsRow(originRowIndex))
                        .Select(r => r with { StartRow = enumrowstart, EndRow = enumrowend })
                        .ToList();

                    sqref?.Value = string.Join(" ", ranges.Select(r => $"{CellReferenceConverter.GetAlphabeticalIndex(r.StartColumn)}{r.StartRow}:{CellReferenceConverter.GetAlphabeticalIndex(r.EndColumn)}{r.EndRow}"));
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
                    .Append($@" r=""{newRowIndex}"">")
                    .Append(innerXml)
                    .Replace("{{$rowindex}}", newRowIndex.ToString())
                    .Replace("{{$enumrowstart}}", enumrowstart.ToString())
                    .Replace("{{$enumrowend}}", enumrowend.ToString())
                    .Append($"</{fullPrefix}{row.Name.LocalName}>");

                ProcessFormulas(rowXml, newRowIndex);
                await writer.WriteRawAsync(CleanXml(rowXml, prefix).ToString()).ConfigureAwait(false);

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

        await writer.WriteEndElementAsync().ConfigureAwait(false);

        // ECMA-376 element order: sheetData → autoFilter → mergeCells → phoneticPr → conditionalFormatting
        // 1. autoFilter (must come before mergeCells)
        if (!string.IsNullOrEmpty(autoFilterXml))
        {
            await writer.WriteRawAsync(CleanXml(autoFilterXml, prefix)).ConfigureAwait(false);
        }

        // 2. mergeCells
        if (_newXMergeCellInfos.Count != 0)
        {
            await writer.WriteRawAsync($"<{fullPrefix}mergeCells count=\"{_newXMergeCellInfos.Count}\">").ConfigureAwait(false);
            foreach (var cell in _newXMergeCellInfos)
            {
                await writer.WriteRawAsync(cell.ToXmlString(prefix)).ConfigureAwait(false);
            }
            await writer.WriteRawAsync($"</{fullPrefix}mergeCells>\r\n").ConfigureAwait(false);
        }

        // 3. phoneticPr
        if (!string.IsNullOrEmpty(phoneticPrXml))
        {
            await writer.WriteRawAsync(CleanXml(phoneticPrXml, prefix)).ConfigureAwait(false);
        }

        // 4. conditionalFormatting
        if (newConditionalFormatRanges.Count != 0)
        {
            var nodes = newConditionalFormatRanges.Select(cf => cf.Node?.ToString(SaveOptions.DisableFormatting));
            await writer.WriteRawAsync(CleanXml(string.Join("", nodes), prefix)).ConfigureAwait(false);
        }

        foreach (var afterElement in afterSheetData)
        {
#if NET8_0_OR_GREATER
            await afterElement.WriteToAsync(writer, cancellationToken).ConfigureAwait(false);
#else
            afterElement.WriteTo(writer);
#endif
        }

        await writer.WriteEndElementAsync().ConfigureAwait(false);
    }

    //todo: refactor in a way that needs less parameters
    [CreateSyncVersion]
    private async Task<GenerateCellValuesContext> GenerateCellValuesAsync(
        GenerateCellValuesContext generateCellValuesContext, 
        string? endPrefix,
        XmlWriter writer,
        StringBuilder rowXml, 
        int mergeRowCount, 
        bool isHeaderRow, 
        XRowInfo rowInfo, 
        XElement row, 
        int groupingRowDiff,
        string innerXml,
        StringBuilder outerXmlOpen, 
        XElement rowElement, 
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
        var notFirstRowElement = new XElement(rowElement);
        foreach (var cell in notFirstRowElement.Elements(SpreadsheetNs + "c"))
        {
            // Try <v> first (for t="n"/t="b" cells), then <is><t> (for t="inlineStr" cells)
            if (cell.Element(SpreadsheetNs + "v") is { } vTag)
            {
                if (!NonTemplateRegex.IsMatch(vTag.Value))
                    vTag.Value = string.Empty;
            }
            else
            {
                // Handle inline string cells
                var t = cell.Element(SpreadsheetNs + "is")?.Element(SpreadsheetNs + "t");
                if (t is not null && !NonTemplateRegex.IsMatch(t.Value))
                    t.Value = string.Empty;
            }
        }

        foreach (var item in rowInfo.CellIEnumerableValues)
        {
            iEnumerableIndex++;
            var closingTag = !string.IsNullOrEmpty(endPrefix) ? $"{endPrefix}:" : "";

            rowXml.Clear()
                .Append(outerXmlOpen)
                .Append($@" r=""{newRowIndex}"">")
                .Append(innerXml)
                .Replace("{{$rowindex}}", newRowIndex.ToString())
                .Append($"</{closingTag}{row.Name.LocalName}>");

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
            var dict = item as IDictionary<string, object?>;
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
                        var prop = rowInfo.PropsMap[newLines[0]];
                        value = prop.PropertyInfoOrFieldInfo switch
                        {
                            PropertyInfoOrFieldInfo.PropertyInfo => prop.PropertyInfo.GetValue(item),
                            PropertyInfoOrFieldInfo.FieldInfo => prop.FieldInfo.GetValue(item),
                            _ => string.Empty
                        };
                    }

                    if (EvaluateStatement(value, newLines[1], newLines[2]))
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
#if NET8_0_OR_GREATER
                string MatchDelegate(Match x) => replacements.GetValueOrDefault(x.Groups[1].Value, "");
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
                        if (!dict!.TryGetValue(prop.Key, out cellValue))
                            continue;
                    }
                    else if (rowInfo.IsDataTable)
                    {
                        cellValue = dataRow![prop.Key];
                    }
                    else
                    {
                        cellValue = propInfo.GetValue(item);
                    }

                    if (cellValue is null)
                        continue;

                    var type = isDictOrTable
                        ? prop.Value.UnderlyingMemberType
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

                    if (isHeaderRow && row.Value.Contains(key))
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
            {
                rowIndexDiff += rowInfo.IEnumerableMercell?.Height ?? 1; //TODO:base on the merge size
            }
            else
            {
                innerXml = string.Concat(notFirstRowElement.Nodes());
                isFirst = false;
            }

            var mergeBaseRowIndex = newRowIndex;
            newRowIndex += rowInfo.IEnumerableMercell?.Height ?? 1;

            // Replace {{$rowindex}} in the already-built substXmlRow
            rowXml.Replace("{{$rowindex}}", mergeBaseRowIndex.ToString());

            // replace formulas
            ProcessFormulas(rowXml, newRowIndex);
            var finalXml = CleanXml(rowXml, endPrefix).ToString();
            await writer.WriteRawAsync(finalXml).ConfigureAwait(false);

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

                var newRow = new XElement(row);
                newRow.SetAttributeValue("r", mergeBaseRowIndex.ToString());

                var oldCells = row.Elements(SpreadsheetNs + "c");
                var newCells = newRow.Elements(SpreadsheetNs + "c");

                // all v replace by empty
                // TODO: remove c/v
                foreach (var (newCell, oldCell) in newCells.Zip(oldCells, (x1, x2) => (x1, x2)))
                {
                    newCell.Attribute("t")?.Remove();
                    newCell.RemoveNodes();
                    newCell.Value = oldCell.Value.Replace("{{$rowindex}}", mergeBaseRowIndex.ToString());
                }

                await writer.WriteRawAsync(CleanXml(newRow.ToString(), endPrefix)).ConfigureAwait(false);
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
            .SelectMany(s => s.Row.Elements(SpreadsheetNs + "c"))
            .Where(s => !string.IsNullOrEmpty(s.Value))
            .Select(s =>
            {
                var att = s.Attribute("r");
                return new XChildNode
                {
                    InnerText = s.Value,
                    ColIndex = StringHelper.GetLetters(att.Value),
                    RowIndex = StringHelper.GetNumber(att.Value)
                };
            })
            .OrderBy(x => x.RowIndex)
            .ToList();

        var mergeColumns = columns.Where(s => s.InnerText?.Contains("@merge") is true).ToList();
        var endMergeColumns = columns.Where(s => s.InnerText?.Contains("@endmerge") is true).ToList();
        var mergeLimitColumn = mergeColumns.FirstOrDefault(x => x.InnerText?.Contains("@mergelimit") is true);

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
                x.ColIndex == taggedColumn.Key.ColIndex && 
                x.RowIndex > taggedColumn.Key.RowIndex &&
                x.RowIndex < taggedColumn.Value.RowIndex));
        }

        var lastMergeCellIndexes = new Dictionary<int, MergeCellIndex>();
        foreach (var rowInfo in xRowInfos)
        {
            var row = rowInfo.Row;
            var childNodes = row.Elements();

            foreach (var node in childNodes)
            {
                var att = node.Attribute("r");
                var nodeLetter = StringHelper.GetLetters(att.Value);
                var nodeNumber = StringHelper.GetNumber(att.Value);

                if (!string.IsNullOrEmpty(node.Value))
                {
                    var xmlNodes = calculatedColumns
                        .Where(j =>
                            j.InnerText == node.Value &&
                            j.ColIndex == nodeLetter)
                        .OrderBy(s => s.RowIndex)
                        .ToList();

                    if (xmlNodes.Count > 1)
                    {
                        if (mergeLimitColumn is not null)
                        {
                            var limitedNode = calculatedColumns.First(j =>
                                j.ColIndex == mergeLimitColumn.ColIndex && j.RowIndex == nodeNumber);

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

                node.SetAttributeValue("r", $"{nodeLetter}{{{{$rowindex}}}}");
            }
        }
    }

    private void ProcessFormulas(StringBuilder rowXml, int rowIndex)
    {
        // exit if no formula is found
        if (!rowXml.ToString().Contains("$="))
            return;

        // adding dummy element for correctly parsing namespace prefix
        rowXml.Insert(0, $"<d xmlns:x14ac=\"{Schemas.SpreadsheetmlXmlX14Ac}\">");
        rowXml.Append("</d>");

        var rowElement = XElement.Parse(rowXml.ToString());

        var index = 1;
        foreach (var cell in rowElement.Descendants(SpreadsheetNs + "c"))
        {
            // convert cells starting with '$=' into formulas
            /* Target:
                 <c r="C8" s="3">
                    <f>SUM(C2:C7)</f>
                </c>
            */
            foreach (var str in cell.Elements(SpreadsheetNs + "is"))
            {
                if (str.Value.StartsWith("$="))
                {
                    var fNode = new XElement(SpreadsheetNs + "f");
                    fNode.SetValue(str.Value[2..]);
                    str.AddBeforeSelf(fNode);
                    str.Remove();

                    var celRef = CellReferenceConverter.GetCellFromCoordinates(index, rowIndex);
                    _calcChainCellRefs.Add(celRef);
                }
            }

            index++;
        }

        rowXml.Clear();
        rowXml.Append(rowElement.FirstNode);
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
    private static string CleanXml(string? xml, string? prefix) => CleanXml(new StringBuilder(xml), prefix).ToString();
    private static StringBuilder CleanXml(StringBuilder xml, string? prefix = null)
    {
        var sb = xml
            .Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "")
            .Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");

        return !string.IsNullOrEmpty(prefix) 
            ? sb.Replace($"xmlns:{prefix}=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "") 
            : sb;
    }

    private static void InjectSharedStrings(IDictionary<int, string> sharedStrings, IEnumerable<XElement> rows)
    {
        foreach (var row in rows)
        {
            var cells = row.Elements(SpreadsheetNs + "c");
            foreach (var cell in cells)
            {
                var t = cell.Attribute("t");
                var v = cell.Element(SpreadsheetNs + "v");
                
                if (v?.Value is null || t?.Value != "s")
                    continue;

                //needs to check if sharedstring exists or not
                if (sharedStrings is null || !sharedStrings.TryGetValue(int.Parse(v.Value), out var shared))
                    continue;

                // change type = inlineStr and replace its value
                v.Remove();

                var tNode = new XElement(SpreadsheetNs + "t", shared);
                var isNode = new XElement(SpreadsheetNs + "is", tNode);
                cell.Add(isNode);
                cell.SetAttributeValue("t", "inlineStr");
            }
        }
    }

    private static void SetCellType(XElement cell, string type)
    {
        // Force inlineStr for strings
        if (type == "str") 
            type = "inlineStr";

        if (type == "inlineStr")
        {
            // Ensure <is><t>...</t></is>
            cell.SetAttributeValue("t", "inlineStr");

            if (cell.Element(SpreadsheetNs + "v") is { } v)
            {
                var text = v.Value;
                v.Remove();

                var tNode = new XElement(SpreadsheetNs + "t", text);
                var isNode = new XElement(SpreadsheetNs + "is", tNode);

                cell.Add(isNode);
                cell.SetAttributeValue("t", "inlineStr");
            }
            else if (cell.Element(SpreadsheetNs + "is") is null)
            {
                // Create empty <is><t></t></is> if neither <v> nor <is> exists
                var tNode = new XElement(SpreadsheetNs + "t");
                var isNode = new XElement(SpreadsheetNs + "is", tNode);

                cell.Add(isNode);
            }
        }
        else
        {
            // Ensure <v>...</v>
            // For numbers/booleans, we remove 't' attribute to let it be default (number) 
            // or we could set it to 'n' explicitly, but removing is safer for general number types
            if (type == "b")
                cell.SetAttributeValue("t", "b");
            else
                cell.Attribute("t")?.Remove();

            if (cell.Element(SpreadsheetNs + "is") is { } isNode)
            {
                var tNode = isNode.Element(SpreadsheetNs + "t");
                var text = tNode?.Value ?? string.Empty;
                isNode.Remove();

                cell.Add(new XElement(SpreadsheetNs + "v", text));
            }
        }
    }

    private void UpdateDimensionAndGetRowsInfo(IDictionary<string, object?> inputMaps, XElement worksheet, IEnumerable<XElement> rows, bool changeRowIndex = true)
    {
        string[] refs;
        if (worksheet.Element(SpreadsheetNs + "dimension") is { } dimension && 
            dimension.Attribute("ref") is { Value: var @ref })
        {
            refs = @ref.Split(':');
        }
        else
        {
            // ==== add dimension element if not found ====

            var firstCell = rows.FirstOrDefault()?.Elements(SpreadsheetNs + "c").FirstOrDefault();
            var lastCell = rows.LastOrDefault()?.Elements(SpreadsheetNs + "c").LastOrDefault();

            var dimStart = firstCell?.Attribute("r")?.Value ?? "";
            var dimEnd = lastCell?.Attribute("r")?.Value ?? "";

            refs = [dimStart, dimEnd];

            dimension = new XElement(SpreadsheetNs + "dimension");
            worksheet.AddFirst(dimension);
        }

        var maxRowIndexDiff = 0;
        foreach (var row in rows)
        {
            // ==== get ienumerable infomation & maxrowindexdiff ====

            var xRowInfo = new XRowInfo { Row = row };
            _xRowInfos.Add(xRowInfo);

            foreach (var cell in row.Elements(SpreadsheetNs + "c"))
            {
                var r = cell.Attribute("r")?.Value;

                // ==== mergecells ====
                if (_xMergeCellInfos.TryGetValue(r, out var merCell))
                {
                    xRowInfo.RowMercells ??= [];
                    xRowInfo.RowMercells.Add(merCell);
                }

                if (changeRowIndex)
                {
                    cell.SetAttributeValue("r", $"{StringHelper.GetLetters(r)}{{{{$rowindex}}}}");
                }

                var v = cell.Element(SpreadsheetNs + "v") ?? cell.Element(SpreadsheetNs + "is")?.Element(SpreadsheetNs + "t");
                if (v?.Value is null)
                    continue;

                var matches = IsExpressionRegex.Matches(v.Value)
                    .Cast<Match>()
                    .Select(x => x.Value)
                    .Distinct()
                    .ToArray();

                var matchCount = matches.Length;
                var isMultiMatch = matchCount > 1 || (matchCount == 1 && v.Value != $"{{{{{matches[0]}}}}}");

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

                        v?.Value = v.Value.Replace($"{{{{{propNames[0]}}}}}", "");
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
                                    xRowInfo.IEnumerablePropName = propNames[0];
                                    xRowInfo.IEnumerableGenericType = element.GetType();

                                    if (element is IDictionary<string, object?> dic)
                                    {
                                        xRowInfo.IsDictionary = true;
                                        xRowInfo.PropsMap = dic.ToDictionary(
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
                                                var propInfo = new MemberInfo
                                                {
                                                    FieldInfo = f,
                                                    PropertyInfoOrFieldInfo = PropertyInfoOrFieldInfo.FieldInfo,
                                                    UnderlyingMemberType = Nullable.GetUnderlyingType(f.FieldType) ?? f.FieldType
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
                            v.Value = v.Value.Replace($"{{{{{propNames[0]}}}}}", propNames[1]);
                            break;
                        }
                        if (!xRowInfo.PropsMap.TryGetValue(propNames[1], out var prop))
                        {
                            v?.Value = v.Value.Replace($"{{{{{propNames[0]}.{propNames[1]}}}}}", "");
                            continue;
                        }
                        // auto check type https://github.com/mini-software/MiniExcel/issues/177
                        var type = prop.UnderlyingMemberType; //avoid nullable

                        if (isMultiMatch)
                        {
                            SetCellType(cell, "str");
                        }
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
                        {
                            SetCellType(cell, "n");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                        {
                            SetCellType(cell, "b");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                        {
                            SetCellType(cell, "str");
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
                                col => new MemberInfo { UnderlyingMemberType = Nullable.GetUnderlyingType(col.DataType) }
                            );
                        }

                        var column = dt.Columns[propNames[1]];
                        var type = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType; //avoid nullable
                        if (!xRowInfo.PropsMap.ContainsKey(propNames[1]))
                            throw new InvalidDataException($"{propNames[0]} doesn't have {propNames[1]} property");

                        if (isMultiMatch)
                        {
                            SetCellType(cell, "str");
                        }
                        else if (TypeHelper.IsNumericType(type) && !type.IsEnum)
                        {
                            SetCellType(cell, "n");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.Boolean)
                        {
                            SetCellType(cell, "b");
                        }
                        else if (Type.GetTypeCode(type) == TypeCode.DateTime)
                        {
                            SetCellType(cell, "str");
                        }
                    }
                    else
                    {
                        var cellValueStr = cellValue?.ToString(); // value did encodexml, so don't duplicate encode value (https://gitee.com/dotnetchina/MiniExcel/issues/I4DQUN)
                        if (isMultiMatch || cellValue is string) // if matchs count over 1 need to set type=str (https://user-images.githubusercontent.com/12729184/114530109-39d46d00-9c7d-11eb-8f6b-52ad8600aca3.png)
                        {
                            SetCellType(cell, "str");
                        }
                        else if (decimal.TryParse(cellValueStr, out var outV))
                        {
                            SetCellType(cell, "n");
                            cellValueStr = outV.ToString(CultureInfo.InvariantCulture);
                        }
                        else if (cellValue is bool b)
                        {
                            SetCellType(cell, "b");
                            cellValueStr = b ? "1" : "0";
                        }
                        else if (cellValue is DateTime timestamp)
                        {
                            //c.SetAttribute("t", "d");
                            cellValueStr = timestamp.ToString("yyyy-MM-dd HH:mm:ss");
                        }

                        if (string.IsNullOrEmpty(cellValueStr) && string.IsNullOrEmpty(cell.Attribute("t")?.Value))
                        {
                            SetCellType(cell, "str");
                        }

                        // Re-acquire v after SetCellType may have changed DOM structure
                        v = cell.Element(SpreadsheetNs + "v") ?? cell.Element(SpreadsheetNs + "is")?.Element(SpreadsheetNs + "t");
                        v?.SetValue(v.Value.Replace($"{{{{{propNames[0]}}}}}", cellValueStr)); //TODO: auto check type and set value
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
            dimension.SetAttributeValue("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
        }
        else
        {
            var letter = StringHelper.GetLetters(refs[0]);
            var digit = StringHelper.GetNumber(refs[0]);
            dimension.SetAttributeValue("ref", $"A1:{letter}{digit + maxRowIndexDiff}");
        }
    }

    private static bool EvaluateStatement(object? tagValue, string comparisonOperator, string value)
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
                "<=" => dtg <= doubleNumber,
                _ => throw new InvalidDataException($"Invalid comparison oeprator: {comparisonOperator}")
            },
            
            int itg when int.TryParse(value, out var intNumber) => comparisonOperator switch
            {
                "==" => itg.Equals(intNumber),
                "!=" => !itg.Equals(intNumber),
                ">" => itg > intNumber,
                "<" => itg < intNumber,
                ">=" => itg >= intNumber,
                "<=" => itg <= intNumber,
                _ => throw new InvalidDataException($"Invalid comparison oeprator: {comparisonOperator}")
            },
            
            DateTime dttg when DateTime.TryParse(value, out var date) => comparisonOperator switch
            {
                "==" => dttg.Equals(date),
                "!=" => !dttg.Equals(date),
                ">" => dttg > date,
                "<" => dttg < date,
                ">=" => dttg >= date,
                "<=" => dttg <= date,
                _ => throw new InvalidDataException($"Invalid comparison oeprator: {comparisonOperator}")
            },
            
            string stg => comparisonOperator switch
            {
                "==" => stg == value,
                "!=" => stg != value,
                _ => throw new InvalidDataException($"Invalid comparison oeprator: {comparisonOperator}")
            },
            
            _ => false
        };
    }
}
