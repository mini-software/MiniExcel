using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlSheetReader : IExcelReader
    {
        private const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private List<SheetRecord> _sheetRecords = null;
        private Dictionary<int, string> _SharedStrings;
        private ExcelOpenXmlStyles _style;

        internal Dictionary<int, string> GetSharedStrings(ExcelOpenXmlZip archive)
        {
            var sharedStringsEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry == null)
                return null;
            using (var stream = sharedStringsEntry.Open())
            {
                var xl = XElement.Load(stream);
                var ts = xl.Descendants(ExcelOpenXmlXName.T).Select((s, i) => new { i, v = s.Value?.ToString() })
                      .ToDictionary(s => s.i, s => s.v)
                ;//TODO:need recode
                return ts;
            }
        }

        internal IEnumerable<SheetRecord> ReadWorkbook(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            using (var stream = entries.Single(w => w.FullName == "xl/workbook.xml").Open())
            using (XmlReader reader = XmlReader.Create(stream, XmlSettings))
            {
                if (!reader.IsStartElement("workbook", ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheets", ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("sheet", ns))
                            {
                                yield return new SheetRecord(
                                    reader.GetAttribute("name"),
                                    uint.Parse(reader.GetAttribute("sheetId")),
                                    reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
                                );
                                reader.Skip();
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        yield break;
                    }
                }
            }
        }

        internal void ReadWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            _sheetRecords = ReadWorkbook(entries).ToList();
            //_styles = ReadStyle(entries).ToList();

            using (var stream = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels").Open())
            using (XmlReader reader = XmlReader.Create(stream, XmlSettings))
            {
                if (!reader.IsStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    return;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    {
                        string rid = reader.GetAttribute("Id");
                        foreach (var sheet in _sheetRecords)
                        {
                            if (sheet.Rid == rid)
                            {
                                sheet.Path = reader.GetAttribute("Target");
                                break;
                            }
                        }

                        reader.Skip();
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }

        public IEnumerable<IDictionary<string, object>> Query(Stream stream, bool UseHeaderRow = false)
        {
            using (var archive = new ExcelOpenXmlZip(stream))
            {
                //TODO:need to optimize
                _SharedStrings = GetSharedStrings(archive);


                // if sheets count > 1 need to read xl/_rels/workbook.xml.rels and 
                var sheets = archive.Entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                    || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                );
                ZipArchiveEntry firstSheetEntry = null;
                if (sheets.Count() > 1)
                {
                    ReadWorkbookRels(archive.Entries);
                    firstSheetEntry = sheets.Single(w => w.FullName == $"xl/{_sheetRecords[0].Path}" || w.FullName == $"/xl/{_sheetRecords[0].Path}");
                }
                else
                    firstSheetEntry = sheets.Single();


                // TODO: need to optimize performance
                var withoutCR = false;

                var maxRowIndex = -1;
                var maxColumnIndex = -1;

                //TODO: merge one open read
                using (var firstSheetEntryStream = firstSheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(firstSheetEntryStream, XmlSettings))
                {
                    while (reader.Read())
                    {
                        if (reader.IsStartElement("c", ns))
                        {
                            var r = reader.GetAttribute("r");
                            if (r != null)
                            {
                                if (ReferenceHelper.ParseReference(r, out var column, out var row))
                                {
                                    column = column - 1;
                                    row = row - 1;
                                    maxRowIndex = Math.Max(maxRowIndex, row);
                                    maxColumnIndex = Math.Max(maxColumnIndex, column);
                                }
                            }
                            else
                            {
                                withoutCR = true;
                                break;
                            }
                        }
                        //this method logic depends on dimension to get maxcolumnIndex, if without dimension then it need to foreach all rows first time to get maxColumn and maxRowColumn
                        else if (reader.IsStartElement("dimension", ns))
                        {
                            var @ref = reader.GetAttribute("ref");
                            if (string.IsNullOrEmpty(@ref))
                                throw new InvalidOperationException("Without sheet dimension data");
                            var rs = @ref.Split(':');
                            // issue : https://github.com/shps951023/MiniExcel/issues/102
                            if (ReferenceHelper.ParseReference(rs.Length == 2 ? rs[1] : rs[0], out int cIndex, out int rIndex))
                            {
                                maxColumnIndex = cIndex - 1;
                                maxRowIndex = rIndex - 1;
                                break;
                            }
                            else
                                throw new InvalidOperationException("Invaild sheet dimension start data");
                        }
                    }
                }

                if (withoutCR)
                {
                    using (var firstSheetEntryStream = firstSheetEntry.Open())
                    using (XmlReader reader = XmlReader.Create(firstSheetEntryStream, XmlSettings))
                    {
                        if (!reader.IsStartElement("worksheet", ns))
                            yield break;
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            yield break;
                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("sheetData", ns))
                            {
                                if (!XmlReaderHelper.ReadFirstContent(reader))
                                    continue;

                                while (!reader.EOF)
                                {
                                    if (reader.IsStartElement("row", ns))
                                    {
                                        maxRowIndex++;

                                        if (!XmlReaderHelper.ReadFirstContent(reader))
                                            continue;

                                        //Cells
                                        {
                                            var cellIndex = -1;
                                            while (!reader.EOF)
                                            {
                                                if (reader.IsStartElement("c", ns))
                                                {
                                                    cellIndex++;
                                                    maxColumnIndex = Math.Max(maxColumnIndex, cellIndex);
                                                }


                                                if (!XmlReaderHelper.SkipContent(reader))
                                                    break;
                                            }
                                        }
                                    }
                                    else if (!XmlReaderHelper.SkipContent(reader))
                                    {
                                        break;
                                    }
                                }
                            }
                            else if (!XmlReaderHelper.SkipContent(reader))
                            {
                                break;
                            }
                        }

                    }
                }


                using (var firstSheetEntryStream = firstSheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(firstSheetEntryStream, XmlSettings))
                {
                    if (!reader.IsStartElement("worksheet", ns))
                        yield break;

                    if (!XmlReaderHelper.ReadFirstContent(reader))
                        yield break;

                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("sheetData", ns))
                        {
                            if (!XmlReaderHelper.ReadFirstContent(reader))
                                continue;

                            Dictionary<int, string> headRows = new Dictionary<int, string>();
                            int rowIndex = -1;
                            int nextRowIndex = 0;
                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("row", ns))
                                {
                                    nextRowIndex = rowIndex + 1;
                                    if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                                        rowIndex = arValue - 1; // The row attribute is 1-based
                                    else
                                        rowIndex++;
                                    if (!XmlReaderHelper.ReadFirstContent(reader))
                                        continue;

                                    // fill empty rows
                                    {
                                        if (nextRowIndex < rowIndex)
                                        {
                                            for (int i = nextRowIndex; i < rowIndex; i++)
                                                if (UseHeaderRow)
                                                    yield return Helpers.GetEmptyExpandoObject(headRows);
                                                else
                                                    yield return Helpers.GetEmptyExpandoObject(maxColumnIndex);
                                        }
                                    }

                                    // Set Cells
                                    {
                                        var cell = UseHeaderRow ? Helpers.GetEmptyExpandoObject(headRows) : Helpers.GetEmptyExpandoObject(maxColumnIndex);
                                        var columnIndex = withoutCR ? -1 : 0;
                                        while (!reader.EOF)
                                        {
                                            if (reader.IsStartElement("c", ns))
                                            {
                                                var aS = reader.GetAttribute("s");
                                                var cellValue = ReadCell(reader, columnIndex, withoutCR, out var _columnIndex);
                                                columnIndex = _columnIndex;

                                                // TODO: bad code smell 
                                                if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                                                {
                                                    int xfIndex = -1;
                                                    if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                                                    {
                                                        xfIndex = styleIndex;
                                                    }
                                                    // only when have s attribute then load styles xml data
                                                    if (_style == null)
                                                        _style = new ExcelOpenXmlStyles(archive);
                                                    //if not using First Head then using 1,2,3 as index
                                                    if (UseHeaderRow)
                                                    {
                                                        if (rowIndex == 0)
                                                        {
                                                            var customStyleCellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue)?.ToString();
                                                            if (!string.IsNullOrWhiteSpace(customStyleCellValue))
                                                                headRows.Add(columnIndex, customStyleCellValue);
                                                        }
                                                        else
                                                        {
                                                            if (headRows.ContainsKey(columnIndex))
                                                            {
                                                                var key = headRows[columnIndex];
                                                                var v = _style.ConvertValueByStyleFormat(int.Parse(aS), cellValue);
                                                                cell[key] = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //if not using First Head then using A,B,C as index
                                                        cell[Helpers.GetAlphabetColumnName(columnIndex)] = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                                    }
                                                }
                                                else
                                                {
                                                    if (UseHeaderRow)
                                                    {
                                                        if (rowIndex == 0)
                                                        {
                                                            var valueString = cellValue?.ToString();
                                                            if (!string.IsNullOrWhiteSpace(valueString))
                                                                headRows.Add(columnIndex, valueString);
                                                        }
                                                        else
                                                        {
                                                            if (headRows.ContainsKey(columnIndex))
                                                                cell[headRows[columnIndex]] = cellValue;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //if not using First Head then using A,B,C as index
                                                        cell[Helpers.GetAlphabetColumnName(columnIndex)] = cellValue;
                                                    }
                                                }
                                            }
                                            else if (!XmlReaderHelper.SkipContent(reader))
                                                break;
                                        }

                                        if (UseHeaderRow && rowIndex == 0)
                                            continue;

                                        yield return cell;
                                    }
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }

                        }
                        else if (!XmlReaderHelper.SkipContent(reader))
                        {
                            break;
                        }
                    }
                }
            }
        }

        public IEnumerable<T> Query<T>(Stream stream) where T : class, new()
        {
            var type = typeof(T);
            var props = Helpers.GetExcelCustomPropertyInfos(type);
            foreach (var item in Query(stream, true))
            {
                var v = new T();
                foreach (var pInfo in props)
                {
                    if (item.ContainsKey(pInfo.ExcelColumnName))
                    {
                        object newV = null;
                        object itemValue = item[pInfo.ExcelColumnName];

                        if (itemValue == null)
                            continue;

                        if (pInfo.ExcludeNullableType == typeof(Guid))
                        {
                            newV = Guid.Parse(itemValue.ToString());
                        }
                        else if (pInfo.ExcludeNullableType == typeof(DateTime))
                        {
                            var vs = itemValue.ToString();
                            if (DateTime.TryParse(vs, out var _v))
                                newV = _v;
                            else if (DateTime.TryParseExact(vs, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out var _v2))
                                newV = _v2;
                            else if (double.TryParse(vs, out var _d))
                                newV = DateTimeHelper.FromOADate(_d);
                            else
                                throw new InvalidCastException($"{vs} can't cast to datetime");
                        }
                        else if (pInfo.ExcludeNullableType == typeof(bool))
                        {
                            var vs = itemValue.ToString();
                            if (vs == "1")
                                newV = true;
                            else if (vs == "0")
                                newV = false;
                            else
                                newV = bool.Parse(vs);
                        }
                        // solve : https://github.com/shps951023/MiniExcel/issues/138
                        else
                            newV = Convert.ChangeType(itemValue, pInfo.ExcludeNullableType);
                        pInfo.Property.SetValue(v, newV);
                    }
                }
                yield return v;
            }
        }

        private object ReadCell(XmlReader reader, int nextColumnIndex, bool withoutCR, out int columnIndex)
        {
            int xfIndex = -1;
            var aT = reader.GetAttribute("t");
            var aR = reader.GetAttribute("r");

            if (withoutCR)
                columnIndex = nextColumnIndex + 1;
            //TODO:need to check only need nextColumnIndex or columnIndex
            else if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                columnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                columnIndex = nextColumnIndex;

            if (!XmlReaderHelper.ReadFirstContent(reader))
                return null;

            object value = null;
            while (!reader.EOF)
            {
                if (reader.IsStartElement("v", ns))
                {
                    string rawValue = reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (reader.IsStartElement("is", ns))
                {
                    string rawValue = StringHelper.ReadStringItem(reader);
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (!XmlReaderHelper.SkipContent(reader))
                {
                    break;
                }
            }

            return value;
        }

        private void ConvertCellValue(string rawValue, string aT, int xfIndex, out object value)
        {
            const NumberStyles style = NumberStyles.Any;
            var invariantCulture = CultureInfo.InvariantCulture;

            switch (aT)
            {
                case "s": //// if string
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        if (_SharedStrings.ContainsKey(sstIndex))
                            value = _SharedStrings[sstIndex];
                        else
                            value = sstIndex;
                        return;
                    }

                    value = rawValue;
                    return;
                case "inlineStr": //// if string inline
                case "str": //// if cached formula string
                    value = Helpers.ConvertEscapeChars(rawValue);
                    return;
                case "b": //// boolean
                    value = rawValue == "1";
                    return;
                case "d": //// ISO 8601 date
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;
                case "e": //// error
                    value = rawValue;
                    return;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out double number))
                    {
                        value = number;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = null,
        };
    }
}
