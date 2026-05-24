using static MiniExcelLib.Core.Helpers.ImageHelper;

namespace MiniExcelLib.OpenXml.Writer;

internal partial class OpenXmlWriter
{
    private const string DefaultCellStyleIndex = "0";
    private const string HeaderCellStyleIndex = "1";
    private const string RegularCellStyleIndex = "2";
    private const string DateCellStyleIndex = "3";
    private const string FillCellStyleIndex = "4";
    private const string TimeCellStyleIndex = "5";
    
    private static readonly DateTime ExcelZeroDate = new(1899, 12, 31);
    
    private readonly Dictionary<string, string> _zipContentsMap = [];
    private readonly Dictionary<string, int> _sharedStrings = [];
    
    private IEnumerable<(SheetDto Sheet, object? Data)> GetSheets()
    {
        var sheetId = 0;
        if (_value is IDictionary<string, object?> dictionary)
        {
            foreach (var sheet in dictionary)
            {
                ThrowHelper.ThrowIfInvalidSheetName(sheet.Key);
                
                sheetId++;
                var sheetInfos = GetSheetInfos(sheet.Key);
                yield return (sheetInfos.ToDto(sheetId), sheet.Value);
            }

            yield break;
        }

        if (_value is DataSet dataSet)
        {
            foreach (DataTable dt in dataSet.Tables)
            {
                sheetId++;
                var sheetInfos = GetSheetInfos(dt.TableName);
                yield return (sheetInfos.ToDto(sheetId), dt);
            }

            yield break;
        }

        sheetId++;
        var sheetInfo = GetSheetInfos(_sheetName);
        yield return (sheetInfo.ToDto(sheetId), _value);
    }

    private ExcelSheetInfo GetSheetInfos(string sheetName)
    {
        var info = new ExcelSheetInfo
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
        if (_configuration is { FreezeRowCount: <= 0, FreezeColumnCount: <= 0, RightToLeft: false })
            return string.Empty;

        var sb = new StringBuilder();

        // start sheetViews
        sb.Append(WorksheetXml.StartSheetViews);
        sb.Append(WorksheetXml.StartSheetView(rightToLeft: _configuration.RightToLeft));

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

        var activePane = (_configuration.FreezeColumnCount > 0) switch
        {
            true when _configuration.FreezeRowCount > 0 => "bottomRight",
            true => "topRight",
            _ => "bottomLeft"
        };

        var startPane =  WorksheetXml.StartPane(
            xSplit: _configuration.FreezeColumnCount > 0 ? _configuration.FreezeColumnCount : null,
            ySplit: _configuration.FreezeRowCount > 0 ? _configuration.FreezeRowCount : null,
            topLeftCell: CellReferenceConverter.GetCellFromCoordinates(
                _configuration.FreezeColumnCount + 1,
                _configuration.FreezeRowCount + 1
            ),
            activePane: activePane,
            state: "frozen"
        );
        sb.Append(startPane);

        // write pane selections
        if (_configuration is { FreezeColumnCount: > 0, FreezeRowCount: > 0 })
        {
            // freeze row and column
            /*
             <selection pane="topRight" activeCell="B1" sqref="B1"/>
             <selection pane="bottomLeft" activeCell="A3" sqref="A3"/>
             <selection pane="bottomRight" activeCell="B3" sqref="B3"/>
             */
            var cellTr = CellReferenceConverter.GetCellFromCoordinates(_configuration.FreezeColumnCount + 1, 1);
            sb.Append(WorksheetXml.PaneSelection("topRight", cellTr, cellTr));

            var cellBl = CellReferenceConverter.GetCellFromCoordinates(1, _configuration.FreezeRowCount + 1);
            sb.Append(WorksheetXml.PaneSelection("bottomLeft", cellBl, cellBl));

            var cellBr = CellReferenceConverter.GetCellFromCoordinates(_configuration.FreezeColumnCount + 1, _configuration.FreezeRowCount + 1);
            sb.Append(WorksheetXml.PaneSelection("bottomRight", cellBr, cellBr));
        }
        else if (_configuration.FreezeColumnCount > 0)
        {
            // freeze column
            /*
               <selection pane="topRight" activeCell="A1" sqref="A1"/>
            */
            var cellTr = CellReferenceConverter.GetCellFromCoordinates(_configuration.FreezeColumnCount, 1);
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

    private (string StyleIndex, string DataType, string? CellValue) GetCellValue(int rowIndex, int cellIndex, object value, MiniExcelColumnMapping? columnMapping, bool valueIsNull)
    {
        if (valueIsNull)
            return (RegularCellStyleIndex, GetStringType(), string.Empty);

        if (value is string str)
        {
            var styleIndex = columnMapping?.ExcelFormatId is { } fmt and not -1 ? fmt.ToString() : RegularCellStyleIndex;
            return (styleIndex, GetStringType(), str);
        }

        var type = GetValueType(value, columnMapping);

        if (columnMapping is { ExcelFormat: not null, ExcelFormatId: -1 } && value is IFormattable formattableValue)
        {
            var formattedStr = formattableValue.ToString(columnMapping.ExcelFormat, _configuration.Culture);
            return (RegularCellStyleIndex, GetStringType(), formattedStr);
        }

        if (type == typeof(DateTime))
            return GetDateTimeValue((DateTime)value, columnMapping);

        if (type == typeof(DateTimeOffset))
            return GetDateTimeValue(((DateTimeOffset)value).DateTime, columnMapping);

        if (type == typeof(TimeSpan))
            return GetTimeSpanValue((TimeSpan)value, columnMapping);

#if NET
        if (type == typeof(DateOnly))
            return GetDateTimeValue(((DateOnly)value).ToDateTime(default), columnMapping);

        if (type == typeof(TimeOnly))
            return GetTimeSpanValue(((TimeOnly)value).ToTimeSpan(), columnMapping);
#endif

        if (type.IsEnum)
        {
            string? description = null;
            var name = Enum.GetName(type, value);

            if (!string.IsNullOrEmpty(name))
            {
                var descAttr = type.GetField(name)?.GetCustomAttribute<DescriptionAttribute>();
                description = descAttr?.Description ?? name;
            }

            description ??= value.ToString();
            return (RegularCellStyleIndex, GetStringType(), description);
        }

        if (TypeHelper.IsNumericType(type))
        {
            var cellValue = GetNumericValue(value, type);
            if (columnMapping?.ExcelFormat is null)
            {
                var dataType = ReferenceEquals(_configuration.Culture, CultureInfo.InvariantCulture) ? ExcelDataTypes.Numeric : GetStringType();
                return (RegularCellStyleIndex, dataType, cellValue);
            }

            return (columnMapping.ExcelFormatId.ToString(), ExcelDataTypes.Numeric, cellValue);
        }

        if (type == typeof(bool))
            return (RegularCellStyleIndex, ExcelDataTypes.Boolean, (bool)value ? "1" : "0");

        if (type == typeof(byte[]) && _configuration.EnableConvertByteArray)
        {
            if (!_configuration.EnableWriteFilePath)
                return (FillCellStyleIndex, ExcelDataTypes.CalculatedString, "");
            
            var base64 = GetFileValue(rowIndex, cellIndex, value);
            return (FillCellStyleIndex, ExcelDataTypes.InlineString, base64);  
        }

        return (RegularCellStyleIndex, GetStringType(), value.ToString());
        
        string GetStringType()
        {
            if (columnMapping?.ExcelColumnType == ColumnType.Formula)
                return ExcelDataTypes.CalculatedString;
            
            return _configuration.StringStorageMode == StringStorageMode.Shared 
                ? ExcelDataTypes.SharedString 
                : ExcelDataTypes.InlineString;
        }
    }

    private static Type GetValueType(object value, MiniExcelColumnMapping? columnInfo)
    {
        if (columnInfo is { Key: null })
            return columnInfo.ExcludeNullableType; //sometime it doesn't need to re-get type like prop

        // TODO: need to optimize
        // Dictionary need to check type every time, so it's slow..
        var type = value.GetType();
        return Nullable.GetUnderlyingType(type) ?? type;
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

        return decimal.Parse(value.ToString()).ToString(_configuration.Culture);
    }

    private string GetFileValue(int rowIndex, int cellIndex, object value)
    {
        var bytes = (byte[])value;

        // TODO: Setting configuration because it might have high cost?
        var format = GetImageFormat(bytes);
        //it can't insert to zip first to avoid cache image to memory
        //because sheet xml is opening.. https://github.com/mini-software/MiniExcel/issues/304#issuecomment-1017031691
        //int rowIndex, int cellIndex
        var file = new FileDto
        {
            Contents = bytes,
            RowIndex = rowIndex,
            CellIndex = cellIndex,
            SheetId = _currentSheetIndex
        };

        if (format != ImageFormat.Unknown)
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

    //todo:reconsider cultureinfo
    private (string, string, string) GetDateTimeValue(DateTime value, MiniExcelColumnMapping? columnMapping)
    {
        string? cellValue;
        if (!ReferenceEquals(_configuration.Culture, CultureInfo.InvariantCulture))
        {
            cellValue = value.ToString(_configuration.Culture);
            return (RegularCellStyleIndex, ExcelDataTypes.CalculatedString, cellValue);
        }

        var oaDate = CorrectDateTimeValue(value);
        cellValue = oaDate.ToString(CultureInfo.InvariantCulture);
        var format = columnMapping?.ExcelFormatId is { } fmt and not -1 ? fmt.ToString() : DateCellStyleIndex;

        return (format, ExcelDataTypes.Numeric, cellValue);
    }

    private static double CorrectDateTimeValue(DateTime value)
    {
        // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
        // to Lotus 1-2-3 decision from 1983...
        // https://github.com/ClosedXML/ClosedXML/blob/develop/ClosedXML/Extensions/DateTimeExtensions.cs#L45
        const int nonExistent1900Feb29SerialDate = 60;

        var oaDate = value.ToOADate();
        if (oaDate <= nonExistent1900Feb29SerialDate)
        {
            oaDate--;
        }

        return oaDate;
    }

    private (string, string, string) GetTimeSpanValue(TimeSpan value, MiniExcelColumnMapping? columnMapping)
    {
        if (value.TotalDays >= 1)
            return GetDateTimeValue(ExcelZeroDate + value, columnMapping);

        var cellValue = value.TotalDays.ToString(CultureInfo.InvariantCulture);
        var format = columnMapping?.ExcelFormatId is { } fmt and not -1 ? fmt.ToString() : TimeCellStyleIndex;

        return (format, ExcelDataTypes.Numeric, cellValue);
    }

    private static string GetDimensionRef(int maxRowIndex, int maxColumnIndex)
    {
        return (maxRowIndex, maxColumnIndex) switch
        {
            (<= 1, 0) => "A1",
            (_, <= 1) => $"A1:A{maxRowIndex}",
            (0, _) => $"A1:{CellReferenceConverter.GetAlphabeticalIndex(maxColumnIndex - 1)}1",
            _ => $"A1:{CellReferenceConverter.GetAlphabeticalIndex(maxColumnIndex - 1)}{maxRowIndex}"
        };
    }

    private string GetDrawingRelationshipXml(int sheetIndex)
    {
        var drawing = new StringBuilder();
        foreach (var image in _files.Where(w => w.IsImage && w.SheetId == sheetIndex))
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
            if (file.IsImage && file.SheetId == sheetIndex)
            {
                drawing.Append(ExcelXml.DrawingXml(file, fileIndex));
            }
        }

        return drawing.ToString();
    }

    private (string WorkbookXml, string WorkbookRelsXml, Dictionary<int, string> SheetRelsXml) GenerateWorkbookXmls()
    {
        var workbookXml = new StringBuilder();
        var workbookRelsXml = new StringBuilder();
        var sheetsRelsXml = new Dictionary<int, string>();

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

        return (workbookXml.ToString(), workbookRelsXml.ToString(), sheetsRelsXml);
    }

    private string GetContentTypesXml()
    {
        var sb = new StringBuilder(ExcelXml.StartTypes);
        foreach (var p in _zipContentsMap)
        {
            sb.Append(ExcelXml.ContentType(p.Value, p.Key));
        }

        sb.Append(ExcelXml.EndTypes);
        return sb.ToString();
    }
}
