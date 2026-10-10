using MiniExcelLib.OpenXml.Styles;
using XmlReaderHelper = MiniExcelLib.OpenXml.Utils.XmlReaderHelper;

namespace MiniExcelLib.OpenXml.Reader;

internal partial class OpenXmlReader
{
    [CreateSyncVersion]
    internal async IAsyncEnumerable<ExcelCell> QueryCellsAsync(string? sheetName, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var sheetEntry = GetSheetEntry(sheetName);
        var sheetStream = await sheetEntry.OpenAsync(cancellationToken).ConfigureAwait(false);
        await using var disposableSheetStream = sheetStream.ConfigureAwait(false);

        using var reader = XmlReader.Create(sheetStream, XmlReaderHelper.GetXmlReaderSettings());

        if (!reader.IsStartElement("worksheet", Ns))
            yield break;
        if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
            yield break;

        while (!reader.EOF)
        {
            if (!reader.IsStartElement("sheetData", Ns))
            {
                if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                    break;

                continue;
            }

            if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                yield break;

            var rowNumber = 0;
            while (!reader.EOF)
            {
                if (reader.IsStartElement("row", Ns))
                {
                    if (reader.GetAttribute("r") is { } rowAttribute)
                    {
                        if (!int.TryParse(rowAttribute, NumberStyles.None, CultureInfo.InvariantCulture, out rowNumber) ||
                            rowNumber is < 1 or > CellReferenceConverter.MaxRowNumber)
                        {
                            throw new InvalidDataException($"Row index '{rowAttribute}' is outside Excel's valid range.");
                        }
                    }
                    else if (++rowNumber > CellReferenceConverter.MaxRowNumber)
                    {
                        throw new InvalidDataException("The worksheet contains more rows than Excel supports.");
                    }

                    if (!await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
                        continue;

                    var columnNumber = 0;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("c", Ns))
                        {
                            var cell = await ReadCellAsync(reader, rowNumber, columnNumber + 1, cancellationToken).ConfigureAwait(false);
                            columnNumber = cell.ColumnIndex;
                            yield return cell;
                        }
                        else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                        {
                            break;
                        }
                    }
                }
                else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }

            yield break;
        }
    }

    [CreateSyncVersion]
    private async Task<ExcelCell> ReadCellAsync(XmlReader reader, int rowNumber, int nextColumnNumber, CancellationToken cancellationToken)
    {
        var aR = reader.GetAttribute("r");
        var aT = reader.GetAttribute("t");
        var aS = reader.GetAttribute("s");

        var columnNumber = nextColumnNumber;
        if (CellReferenceConverter.TryParseCellReference(aR, out var referenceColumn, out _))
        {
            columnNumber = referenceColumn;
        }
        else if (columnNumber > CellReferenceConverter.MaxColumnNumber)
        {
            throw new InvalidDataException("The worksheet contains more columns than Excel supports.");
        }

        object? value = null;
        var hasFormula = false;
        string? formula = null;

        if (await reader.ReadFirstContentAsync(cancellationToken).ConfigureAwait(false))
        {
            while (!reader.EOF)
            {
                if (reader.IsStartElement("f", Ns))
                {
                    hasFormula = true;
                    var formulaText = await reader.ReadElementContentAsStringAsync()
#if NET
                        .WaitAsync(cancellationToken)
#endif
                        .ConfigureAwait(false);

                    formula = string.IsNullOrEmpty(formulaText) ? null : formulaText;
                }
                else if (reader.IsStartElement("v", Ns))
                {
                    var rawValue = await reader.ReadElementContentAsStringAsync()
#if NET
                        .WaitAsync(cancellationToken)
#endif
                        .ConfigureAwait(false);

                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, -1, out value);
                }
                else if (reader.IsStartElement("is", Ns))
                {
                    var rawValue = await reader.ReadStringItemAsync(cancellationToken).ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, -1, out value);
                }
                else if (!await reader.SkipContentAsync(cancellationToken).ConfigureAwait(false))
                {
                    break;
                }
            }
        }

        // A cell without an s attribute uses the first cell format, so its font is the workbook default font
        var hasStyleIndex = int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var xfIndex);
        ExcelColor? fontColor = null;
        if (hasStyleIndex || Archive.GetEntry(ExcelFileNames.Styles) is not null)
        {
            _style ??= await OpenXmlStyles.CreateAsync(Archive, cancellationToken).ConfigureAwait(false);
            if (hasStyleIndex)
                value = _style.ConvertValueByStyleFormat(xfIndex, value);

            fontColor = _style.GetFontColor(xfIndex);
        }

        var reference = CellReferenceConverter.GetCellFromCoordinates(columnNumber, rowNumber);
        return new ExcelCell(reference, rowNumber, columnNumber, value, hasFormula, formula, fontColor);
    }
}
