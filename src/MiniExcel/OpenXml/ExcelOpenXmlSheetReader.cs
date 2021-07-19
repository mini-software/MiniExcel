using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace MiniExcelLibs.OpenXml
{
    internal class ExcelOpenXmlSheetReader : IExcelReader, IExcelReaderAsync
    {
        private const string _ns = Config.SpreadsheetmlXmlns;
        private List<SheetRecord> _sheetRecords;
        private List<string> _sharedStrings;
        private MergeCells _mergeCells;
        private ExcelOpenXmlStyles _style;
        private ExcelOpenXmlZip _archive;
        private static readonly XmlReaderSettings _xmlSettings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreWhitespace = true,
            XmlResolver = null,
        };

        public ExcelOpenXmlSheetReader(Stream stream)
        {
            _archive = new ExcelOpenXmlZip(stream);
        }

        public IEnumerable<IDictionary<string, object>> Query(bool useHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            var config = (OpenXmlConfiguration)configuration ?? OpenXmlConfiguration.DefaultConfig; //TODO:
            if (!ReferenceHelper.ParseReference(startCell, out var startColumnIndex, out var startRowIndex))
                throw new InvalidDataException($"startCell {startCell} is Invalid");
            startColumnIndex--; startRowIndex--;

            //TODO:need to optimize
            SetSharedStrings();

            // if sheets count > 1 need to read xl/_rels/workbook.xml.rels  
            var sheets = _archive.entries.Where(w => w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
                || w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase)
            );
            ZipArchiveEntry sheetEntry = null;
            if (sheetName != null)
            {
                SetWorkbookRels(_archive.entries);
                var s = _sheetRecords.SingleOrDefault(_ => _.Name == sheetName);
                if (s == null)
                    throw new InvalidOperationException("Please check sheetName/Index is correct");
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || w.FullName == $"/xl/{s.Path}" || w.FullName == s.Path || s.Path == $"/{w.FullName}");
            }
            else if (sheets.Count() > 1)
            {
                SetWorkbookRels(_archive.entries);
                var s = _sheetRecords[0];
                sheetEntry = sheets.Single(w => w.FullName == $"xl/{s.Path}" || w.FullName == $"/xl/{s.Path}");
            }
            else
                sheetEntry = sheets.Single();


            #region MergeCells
            if (config.FillMergedCells)
            {
                _mergeCells = new MergeCells();
                using (var sheetStream = sheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    if (!reader.IsStartElement("worksheet", _ns))
                        yield break;
                    while (reader.Read())
                    {
                        if (reader.IsStartElement("mergeCells", _ns))
                        {
                            if (!XmlReaderHelper.ReadFirstContent(reader))
                                yield break;
                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("mergeCell", _ns))
                                {
                                    var @ref = reader.GetAttribute("ref");
                                    var refs = @ref.Split(':');
                                    if (refs.Length == 1)
                                        continue;

                                    ReferenceHelper.ParseReference(refs[0], out var x1, out var y1);
                                    ReferenceHelper.ParseReference(refs[1], out var x2, out var y2);

                                    _mergeCells.MergesValues.Add(refs[0], null);

                                    // foreach range
                                    var isFirst = true;
                                    for (int x = x1; x <= x2; x++)
                                    {
                                        for (int y = y1; y <= y2; y++)
                                        {
                                            if (!isFirst)
                                                _mergeCells.MergesMap.Add(ReferenceHelper.ConvertXyToCell(x, y), refs[0]);
                                            isFirst = false;
                                        }
                                    }

                                    XmlReaderHelper.SkipContent(reader);
                                }
                                else if (!XmlReaderHelper.SkipContent(reader))
                                {
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            #endregion




            // TODO: need to optimize performance
            var withoutCR = false;
            var maxRowIndex = -1;
            var maxColumnIndex = -1;

            //Q. why need 3 times openstream merge one open read? A. no, zipstream can't use position = 0
            using (var sheetStream = sheetEntry.Open())
            using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
            {
                while (reader.Read())
                {
                    if (reader.IsStartElement("c", _ns))
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
                    else if (reader.IsStartElement("dimension", _ns))
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
                using (var sheetStream = sheetEntry.Open())
                using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
                {
                    if (!reader.IsStartElement("worksheet", _ns))
                        yield break;
                    if (!XmlReaderHelper.ReadFirstContent(reader))
                        yield break;
                    while (!reader.EOF)
                    {
                        if (reader.IsStartElement("sheetData", _ns))
                        {
                            if (!XmlReaderHelper.ReadFirstContent(reader))
                                continue;

                            while (!reader.EOF)
                            {
                                if (reader.IsStartElement("row", _ns))
                                {
                                    maxRowIndex++;

                                    if (!XmlReaderHelper.ReadFirstContent(reader))
                                        continue;

                                    //Cells
                                    {
                                        var cellIndex = -1;
                                        while (!reader.EOF)
                                        {
                                            if (reader.IsStartElement("c", _ns))
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




            using (var sheetStream = sheetEntry.Open())
            using (XmlReader reader = XmlReader.Create(sheetStream, _xmlSettings))
            {
                if (!reader.IsStartElement("worksheet", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheetData", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        Dictionary<int, string> headRows = new Dictionary<int, string>();
                        int rowIndex = -1;
                        int nextRowIndex = 0;
                        bool isFirstRow = true;
                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("row", _ns))
                            {
                                nextRowIndex = rowIndex + 1;
                                if (int.TryParse(reader.GetAttribute("r"), out int arValue))
                                    rowIndex = arValue - 1; // The row attribute is 1-based
                                else
                                    rowIndex++;

                                // row -> c
                                if (!XmlReaderHelper.ReadFirstContent(reader))
                                    continue;

                                // startcell pass rows
                                if (rowIndex < startRowIndex)
                                {
                                    XmlReaderHelper.SkipToNextSameLevelDom(reader);
                                    continue;
                                }


                                // fill empty rows
                                if (!(nextRowIndex < startRowIndex))
                                {
                                    if (nextRowIndex < rowIndex)
                                    {
                                        for (int i = nextRowIndex; i < rowIndex; i++)
                                        {
                                            yield return GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                        }
                                    }
                                }

                                // Set Cells
                                {
                                    var cell = GetCell(useHeaderRow, maxColumnIndex, headRows, startColumnIndex);
                                    var columnIndex = withoutCR ? -1 : 0;
                                    while (!reader.EOF)
                                    {
                                        if (reader.IsStartElement("c", _ns))
                                        {
                                            var aS = reader.GetAttribute("s");
                                            var aR = reader.GetAttribute("r");
                                            var aT = reader.GetAttribute("t");
                                            var cellValue = ReadCellAndSetColumnIndex(reader, ref columnIndex, withoutCR, startColumnIndex, aR, aT);

                                            if (config.FillMergedCells)
                                            {
                                                if (_mergeCells.MergesValues.ContainsKey(aR))
                                                {
                                                    _mergeCells.MergesValues[aR] = cellValue;
                                                }
                                                else if (_mergeCells.MergesMap.ContainsKey(aR))
                                                {
                                                    var mergeKey = _mergeCells.MergesMap[aR];
                                                    object mergeValue = null;
                                                    if(_mergeCells.MergesValues.ContainsKey(mergeKey))
                                                        mergeValue = _mergeCells.MergesValues[mergeKey];
                                                    cellValue = mergeValue;
                                                }
                                            }

                                            if (columnIndex < startColumnIndex)
                                                continue;

                                            if (!string.IsNullOrEmpty(aS)) // if c with s meaning is custom style need to check type by xl/style.xml
                                            {
                                                int xfIndex = -1;
                                                if (int.TryParse(aS, NumberStyles.Any, CultureInfo.InvariantCulture, out var styleIndex))
                                                    xfIndex = styleIndex;

                                                // only when have s attribute then load styles xml data
                                                if (_style == null)
                                                    _style = new ExcelOpenXmlStyles(_archive);
      
                                                cellValue = _style.ConvertValueByStyleFormat(xfIndex, cellValue);
                                                SetCellsValueAndHeaders(cellValue, useHeaderRow, ref headRows, ref isFirstRow, ref cell, columnIndex);
                                            }
                                            else
                                            {
                                                SetCellsValueAndHeaders(cellValue, useHeaderRow, ref headRows, ref isFirstRow, ref cell, columnIndex);
                                            }
                                        }
                                        else if (!XmlReaderHelper.SkipContent(reader))
                                            break;
                                    }

                                    if (isFirstRow)
                                    {
                                        isFirstRow = false; // for startcell logic
                                        if (useHeaderRow)
                                            continue;
                                    }


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

        private static IDictionary<string, object> GetCell(bool useHeaderRow, int maxColumnIndex, Dictionary<int, string> headRows, int startColumnIndex)
        {
            return useHeaderRow ? CustomPropertyHelper.GetEmptyExpandoObject(headRows) : CustomPropertyHelper.GetEmptyExpandoObject(maxColumnIndex, startColumnIndex);
        }

        private void SetCellsValueAndHeaders(object cellValue, bool useHeaderRow, ref Dictionary<int, string> headRows, ref bool isFirstRow, ref IDictionary<string, object> cell, int columnIndex)
        {
            if (useHeaderRow)
            {
                if (isFirstRow) // for startcell logic
                {
                    var cellValueString = cellValue?.ToString();
                    if (!string.IsNullOrWhiteSpace(cellValueString))
                        headRows.Add(columnIndex, cellValueString);
                }
                else
                {
                    if (headRows.ContainsKey(columnIndex))
                    {
                        var key = headRows[columnIndex];
                        cell[key] = cellValue;
                    }
                }
            }
            else
            {
                //if not using First Head then using A,B,C as index
                cell[ColumnHelper.GetAlphabetColumnName(columnIndex)] = cellValue;
            }
        }

        public IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            var type = typeof(T);

            List<ExcelCustomPropertyInfo> props = null;
            var headers = Query(false, sheetName, startCell, configuration).FirstOrDefault()?.Values?.Select(s => s?.ToString())?.ToArray(); //TODO:need to optimize

            var first = true;
            var rowIndex = 0;
            foreach (var item in Query(true, sheetName, startCell, configuration))
            {
                if (first)
                {
                    //TODO: alert don't duplicate column name
                    props = CustomPropertyHelper.GetExcelCustomPropertyInfos(type, headers);
                    first = false;
                }
                var v = new T();
                foreach (var pInfo in props)
                {
                    //TODO:don't need to check every time?
                    if (item.ContainsKey(pInfo.ExcelColumnName))
                    {
                        object newV = null;
                        object itemValue = item[pInfo.ExcelColumnName];

                        if (itemValue == null)
                            continue;

                        newV = TypeHelper.TypeMapping(v, pInfo, newV, itemValue, rowIndex, startCell);
                    }
                }
                rowIndex++;
                yield return v;
            }
        }

        private void SetSharedStrings()
        {
            if (_sharedStrings != null)
                return;
            var sharedStringsEntry = _archive.GetEntry("xl/sharedStrings.xml");
            if (sharedStringsEntry == null)
                return;
            using (var stream = sharedStringsEntry.Open())
            {
                _sharedStrings = GetSharedStrings(stream).ToList();
            }
        }

        internal List<string> GetSharedStrings()
        {
            if (_sharedStrings == null)
                SetSharedStrings();
            return _sharedStrings;
        }

        private IEnumerable<string> GetSharedStrings(Stream stream)
        {
            using (var reader = XmlReader.Create(stream))
            {
                if (!reader.IsStartElement("sst", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("si", _ns))
                    {
                        var value = StringHelper.ReadStringItem(reader);
                        yield return value;
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
        }

        private void SetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            if (_sheetRecords != null)
                return;
            _sheetRecords = GetWorkbookRels(entries);
        }

        internal static IEnumerable<SheetRecord> ReadWorkbook(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            using (var stream = entries.Single(w => w.FullName == "xl/workbook.xml").Open())
            using (XmlReader reader = XmlReader.Create(stream, _xmlSettings))
            {
                if (!reader.IsStartElement("workbook", _ns))
                    yield break;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    yield break;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("sheets", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(reader))
                            continue;

                        while (!reader.EOF)
                        {
                            if (reader.IsStartElement("sheet", _ns))
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

        internal static List<SheetRecord> GetWorkbookRels(ReadOnlyCollection<ZipArchiveEntry> entries)
        {
            var sheetRecords = ReadWorkbook(entries).ToList();

            using (var stream = entries.Single(w => w.FullName == "xl/_rels/workbook.xml.rels").Open())
            using (XmlReader reader = XmlReader.Create(stream, _xmlSettings))
            {
                if (!reader.IsStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    return null;

                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return null;

                while (!reader.EOF)
                {
                    if (reader.IsStartElement("Relationship", "http://schemas.openxmlformats.org/package/2006/relationships"))
                    {
                        string rid = reader.GetAttribute("Id");
                        foreach (var sheet in sheetRecords)
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

            return sheetRecords;
        }

        internal static DataTable QueryAsDataTableImpl(Stream stream, bool useHeaderRow, ref string sheetName, ExcelType excelType, string startCell, IConfiguration configuration)
        {
            if (sheetName == null)
                sheetName = stream.GetSheetNames().First();

            var dt = new DataTable(sheetName);
            var first = true;
            var rows = ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
            foreach (IDictionary<string, object> row in rows)
            {
                if (first)
                {

                    foreach (var key in row.Keys)
                    {
                        var column = new DataColumn(key, typeof(object)) { Caption = key };
                        dt.Columns.Add(column);
                    }

                    dt.BeginLoadData();
                    first = false;
                }

                var newRow = dt.NewRow();
                foreach (var key in row.Keys)
                {
                    newRow[key] = row[key]; //TODO: optimize not using string key
                }

                dt.Rows.Add(newRow);
            }

            dt.EndLoadData();
            return dt;
        }


        private object ReadCellAndSetColumnIndex(XmlReader reader, ref int columnIndex, bool withoutCR, int startColumnIndex, string aR, string aT)
        {
            var newColumnIndex = 0;
            int xfIndex = -1;

            if (withoutCR)
                newColumnIndex = columnIndex + 1;
            //TODO:need to check only need nextColumnIndex or columnIndex
            else if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
                newColumnIndex = referenceColumn - 1; // ParseReference is 1-based
            else
                newColumnIndex = columnIndex;

            columnIndex = newColumnIndex;

            if (columnIndex < startColumnIndex)
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                    return null;
                while (!reader.EOF)
                    if (!XmlReaderHelper.SkipContent(reader))
                        break;
                return null;
            }

            if (!XmlReaderHelper.ReadFirstContent(reader))
                return null;

            object value = null;
            while (!reader.EOF)
            {
                if (reader.IsStartElement("v", _ns))
                {
                    string rawValue = reader.ReadElementContentAsString();
                    if (!string.IsNullOrEmpty(rawValue))
                        ConvertCellValue(rawValue, aT, xfIndex, out value);
                }
                else if (reader.IsStartElement("is", _ns))
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
                case "s":
                    if (int.TryParse(rawValue, style, invariantCulture, out var sstIndex))
                    {
                        if (sstIndex >= 0 && sstIndex < _sharedStrings.Count)
                        {
                            //value = Helpers.ConvertEscapeChars(_SharedStrings[sstIndex]);
                            value = XmlEncoder.DecodeString(_sharedStrings[sstIndex]);
                            return;
                        }
                    }
                    value = null;
                    return;
                case "inlineStr":
                case "str":
                    value = XmlEncoder.DecodeString(rawValue);
                    return;
                case "b":
                    value = rawValue == "1";
                    return;
                case "d":
                    if (DateTime.TryParseExact(rawValue, "yyyy-MM-dd", invariantCulture, DateTimeStyles.AllowLeadingWhite | DateTimeStyles.AllowTrailingWhite, out var date))
                    {
                        value = date;
                        return;
                    }

                    value = rawValue;
                    return;
                case "e":
                    value = rawValue;
                    return;
                default:
                    if (double.TryParse(rawValue, style, invariantCulture, out var n))
                    {
                        value = n;
                        return;
                    }

                    value = rawValue;
                    return;
            }
        }

        public Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, IConfiguration configuration)
        {
            return Task.Run(() => Query(UseHeaderRow, sheetName, startCell, configuration));
        }

        public Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new()
        {
            return Task.Run(() => Query<T>(sheetName, startCell, configuration));
        }
    }
}
