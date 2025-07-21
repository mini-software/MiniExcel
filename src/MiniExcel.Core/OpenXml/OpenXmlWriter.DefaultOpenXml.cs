using MiniExcelLib.Core.Helpers;
using MiniExcelLib.Core.OpenXml.Constants;
using MiniExcelLib.Core.OpenXml.Models;
using MiniExcelLib.Core.OpenXml.Utils;
using MiniExcelLib.Core.OpenXml.Zip;
using MiniExcelLib.Core.Reflection;
using static MiniExcelLib.Core.Helpers.ImageHelper;

namespace MiniExcelLib.Core.OpenXml;

internal partial class OpenXmlWriter : Abstractions.IMiniExcelWriter
{
    private readonly Dictionary<string, ZipPackageInfo> _zipDictionary = [];
    private Dictionary<string, string> _cellXfIdMap;
    
    private IEnumerable<Tuple<SheetDto, object?>> GetSheets()
    {
        var sheetId = 0;
        if (_value is IDictionary<string, object?> dictionary)
        {
            foreach (var sheet in dictionary)
            {
                sheetId++;
                var sheetInfos = GetSheetInfos(sheet.Key);
                yield return Tuple.Create(sheetInfos.ToDto(sheetId), sheet.Value);
            }

            yield break;
        }

        if (_value is DataSet dataSet)
        {
            foreach (DataTable dt in dataSet.Tables)
            {
                sheetId++;
                var sheetInfos = GetSheetInfos(dt.TableName);
                yield return Tuple.Create(sheetInfos.ToDto(sheetId), (object?)dt);
            }

            yield break;
        }

        sheetId++;
        var defaultSheetInfo = GetSheetInfos(_defaultSheetName);
        yield return Tuple.Create(defaultSheetInfo.ToDto(sheetId), _value);
    }

    private ExcellSheetInfo GetSheetInfos(string sheetName)
    {
        var info = new ExcellSheetInfo
        {
            ExcelSheetName = sheetName,
            Key = sheetName,
            ExcelSheetState = SheetState.Visible
        };

        if (_configuration.DynamicSheets is null or [])
            return info;

        var dynamicSheet = _configuration.DynamicSheets.SingleOrDefault(s => s.Key == sheetName);
        if (dynamicSheet is null)
            return info;

        info.ExcelSheetState = dynamicSheet.State;
        if (dynamicSheet.Name is not null) 
            info.ExcelSheetName = dynamicSheet.Name;

        return info;
    }

    private string GetSheetViews()
    {
        // exit early if no style to write
        if (_configuration is { FreezeRowCount: <= 0, FreezeColumnCount: <= 0 })
            return string.Empty;

        var sb = new StringBuilder();

        // start sheetViews
        sb.Append(WorksheetXml.StartSheetViews);
        sb.Append(WorksheetXml.StartSheetView());

        // Write panes
        sb.Append(GetPanes());

        // end sheetViews
        sb.Append(WorksheetXml.EndSheetView);
        sb.Append(WorksheetXml.EndSheetViews);

        return sb.ToString();
    }

    private string GetPanes()
    {
        var sb = new StringBuilder();

        string activePane = (_configuration.FreezeColumnCount > 0) switch
        {
            true when _configuration.FreezeRowCount > 0 => "bottomRight",
            true => "topRight",
            _ => "bottomLeft"
        };
        
        sb.Append(
            WorksheetXml.StartPane(
                xSplit: _configuration.FreezeColumnCount > 0 ? _configuration.FreezeColumnCount : null,
                ySplit: _configuration.FreezeRowCount > 0 ? _configuration.FreezeRowCount : null,
                topLeftCell: ReferenceHelper.ConvertCoordinatesToCell(
                    _configuration.FreezeColumnCount + 1,
                    _configuration.FreezeRowCount + 1
                ),
                activePane: activePane,
                state: "frozen"
            )
        );

        // write pane selections
        if (_configuration is { FreezeColumnCount: > 0, FreezeRowCount: > 0 })
        {
            // freeze row and column
            /*
             <selection pane="topRight" activeCell="B1" sqref="B1"/>
             <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
             <selection pane="bottomRight" activeCell="B3" sqref="B3"/>
             */
            var cellTr = ReferenceHelper.ConvertCoordinatesToCell(_configuration.FreezeColumnCount + 1, 1);
            sb.Append(WorksheetXml.PaneSelection("topRight", cellTr, cellTr));

            var cellBl = ReferenceHelper.ConvertCoordinatesToCell(1, _configuration.FreezeRowCount + 1);
            sb.Append(WorksheetXml.PaneSelection("bottomLeft", cellBl, cellBl));

            var cellBr = ReferenceHelper.ConvertCoordinatesToCell(_configuration.FreezeColumnCount + 1, _configuration.FreezeRowCount + 1);
            sb.Append(WorksheetXml.PaneSelection("bottomRight", cellBr, cellBr));
        }
        else if (_configuration.FreezeColumnCount > 0)
        {
            // freeze column
            /*
               <selection pane="topRight" activeCell="A1" sqref="A1"/>
            */
            var cellTr = ReferenceHelper.ConvertCoordinatesToCell(_configuration.FreezeColumnCount, 1);
            sb.Append(WorksheetXml.PaneSelection("topRight", cellTr, cellTr));

        }
        else
        {
            // freeze row
            /*
                <selection pane="bottomLeft"/>
            */
            sb.Append(WorksheetXml.PaneSelection("bottomLeft", null, null));

        }

        return sb.ToString();
    }

    private Tuple<string, string, string> GetCellValue(int rowIndex, int cellIndex, object value, MiniExcelColumnInfo? columnInfo, bool valueIsNull)
    {
        if (valueIsNull)
            return Tuple.Create("2", "str", string.Empty);

        if (value is string str)
            return Tuple.Create("2", "str", XmlHelper.EncodeXml(str));

        var type = GetValueType(value, columnInfo);

        if (columnInfo is { ExcelFormat: not null, ExcelFormatId: -1 } && value is IFormattable formattableValue)
        {
            var formattedStr = formattableValue.ToString(columnInfo.ExcelFormat, _configuration.Culture);
            return Tuple.Create("2", "str", XmlHelper.EncodeXml(formattedStr));
        }

        if (type == typeof(DateTime))
            return GetDateTimeValue((DateTime)value, columnInfo);

#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly))
            return GetDateTimeValue(((DateOnly)value).ToDateTime(new TimeOnly()), columnInfo);
#endif
        if (type.IsEnum)
        {
            var description = CustomPropertyHelper.DescriptionAttr(type, value);
            return Tuple.Create("2", "str", description ?? value.ToString());
        }

        if (TypeHelper.IsNumericType(type))
        {
            var cellValue = GetNumericValue(value, type);

            if (columnInfo?.ExcelFormat is null)
            {
                var dataType = _configuration.Culture == CultureInfo.InvariantCulture ? "n" : "str";
                return Tuple.Create("2", dataType, cellValue);
            }

            return Tuple.Create(columnInfo.ExcelFormatId.ToString(), (string?)null, cellValue);
        }

        if (type == typeof(bool))
            return Tuple.Create("2", "b", (bool)value ? "1" : "0");

        if (type == typeof(byte[]) && _configuration.EnableConvertByteArray)
        {
            var base64 = GetFileValue(rowIndex, cellIndex, value);
            if (_configuration.EnableWriteFilePath)
            {
                return Tuple.Create("4", "str", XmlHelper.EncodeXml(base64));
            }
            return Tuple.Create("4", "str", "");  
        }

        return Tuple.Create("2", "str", XmlHelper.EncodeXml(value.ToString()));
    }

    private static Type? GetValueType(object value, MiniExcelColumnInfo? columnInfo)
    {
        Type type;
        if (columnInfo is not { Key: null })
        {
            // TODO: need to optimize
            // Dictionary need to check type every time, so it's slow..
            type = value.GetType();
            type = Nullable.GetUnderlyingType(type) ?? type;
        }
        else
        {
            type = columnInfo.ExcludeNullableType; //sometime it doesn't need to re-get type like prop
        }

        return type;
    }

    private string GetNumericValue(object value, Type type)
    {
        if (type.IsAssignableFrom(typeof(decimal)))
            return ((decimal)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(int)))
            return ((int)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(double)))
            return ((double)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(long)))
            return ((long)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(uint)))
            return ((uint)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(ushort)))
            return ((ushort)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(ulong)))
            return ((ulong)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(short)))
            return ((short)value).ToString(_configuration.Culture);

        if (type.IsAssignableFrom(typeof(float)))
            return ((float)value).ToString(_configuration.Culture);

        return (decimal.Parse(value.ToString())).ToString(_configuration.Culture);
    }

    private string GetFileValue(int rowIndex, int cellIndex, object value)
    {
        var bytes = (byte[])value;

        // TODO: Setting configuration because it might have high cost?
        var format = GetImageFormat(bytes);
        //it can't insert to zip first to avoid cache image to memory
        //because sheet xml is opening.. https://github.com/mini-software/MiniExcel/issues/304#issuecomment-1017031691
        //int rowIndex, int cellIndex
        var file = new FileDto()
        {
            Byte = bytes,
            RowIndex = rowIndex,
            CellIndex = cellIndex,
            SheetId = _currentSheetIndex
        };

        if (format != ImageHelper.ImageFormat.Unknown)
        {
            file.Extension = format.ToString();
            file.IsImage = true;
        }
        else
        {
            file.Extension = "bin";
        }

        _files.Add(file);

        //TODO:Convert to base64
        var base64 = $"@@@fileid@@@,{file.Path}";
        return base64;
    }

    private Tuple<string, string?, string> GetDateTimeValue(DateTime value, MiniExcelColumnInfo columnInfo)
    {
        string? cellValue;
        if (!ReferenceEquals(_configuration.Culture, CultureInfo.InvariantCulture))
        {
            cellValue = value.ToString(_configuration.Culture);
            return Tuple.Create("2", (string?)"str", cellValue);
        }

        var oaDate = CorrectDateTimeValue(value);
        cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
        var format = columnInfo?.ExcelFormat is not null ? columnInfo.ExcelFormatId.ToString() : "3";

        return Tuple.Create(format, (string?)null, cellValue);
    }

    private static double CorrectDateTimeValue(DateTime value)
    {
        var oaDate = value.ToOADate();

        // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
        // to Lotus 1-2-3 decision from 1983...
        // https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Extensions/DateTimeExtensions.cs#L45
        const int nonExistent1900Feb29SerialDate = 60;
        if (oaDate <= nonExistent1900Feb29SerialDate)
        {
            oaDate -= 1;
        }

        return oaDate;
    }

    private static string GetDimensionRef(int maxRowIndex, int maxColumnIndex)
    {
        string dimensionRef;
        if (maxRowIndex == 0 && maxColumnIndex == 0)
            dimensionRef = "A1";
        else if (maxRowIndex <= 1 && maxColumnIndex == 0)
            dimensionRef = "A1";
        else if (maxColumnIndex <= 1)
            dimensionRef = $"A1:A{maxRowIndex}";
        else if (maxRowIndex == 0)
            dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}1";
        else
            dimensionRef = $"A1:{ColumnHelper.GetAlphabetColumnName(maxColumnIndex - 1)}{maxRowIndex}";
        return dimensionRef;
    }

    private string GetDrawingRelationshipXml(int sheetIndex)
    {
        var drawing = new StringBuilder();
        foreach (var image in _files.Where(w => w.IsImage && w.SheetId == sheetIndex + 1))
        {
            drawing.AppendLine(ExcelXml.ImageRelationship(image));
        }

        return drawing.ToString();
    }

    private string GetDrawingXml(int sheetIndex)
    {
        var drawing = new StringBuilder();

        for (int fileIndex = 0; fileIndex < _files.Count; fileIndex++)
        {
            var file = _files[fileIndex];
            if (file.IsImage && file.SheetId == sheetIndex + 1)
            {
                drawing.Append(ExcelXml.DrawingXml(file, fileIndex));
            }
        }

        return drawing.ToString();
    }

    private void GenerateWorkBookXmls(
        out StringBuilder workbookXml,
        out StringBuilder workbookRelsXml,
        out Dictionary<int, string> sheetsRelsXml)
    {
        workbookXml = new StringBuilder();
        workbookRelsXml = new StringBuilder();
        sheetsRelsXml = new Dictionary<int, string>();
        var sheetId = 0;
        foreach (var sheetDto in _sheets)
        {
            sheetId++;
            workbookXml.AppendLine(ExcelXml.Sheet(sheetDto, sheetId));

            workbookRelsXml.AppendLine(ExcelXml.WorksheetRelationship(sheetDto));

            //TODO: support multiple drawing
            //TODO: ../drawings/drawing1.xml or /xl/drawings/drawing1.xml
            sheetsRelsXml.Add(sheetDto.SheetIdx, ExcelXml.DrawingRelationship(sheetId));
        }
    }

    private string GetContentTypesXml()
    {
        var sb = new StringBuilder(ExcelXml.StartTypes);
        foreach (var p in _zipDictionary)
        {
            sb.Append(ExcelXml.ContentType(p.Value.ContentType, p.Key));
        }

        sb.Append(ExcelXml.EndTypes);
        return sb.ToString();
    }

    private string GetCellXfId(string styleIndex)
    {
        return _cellXfIdMap.TryGetValue(styleIndex, out var cellXfId) ? cellXfId : styleIndex;
    }
}