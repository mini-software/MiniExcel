using System;
using System.Collections.Generic;
using MiniExcelLibs.OpenXml.Constants;
using MiniExcelLibs.Utils;
using MiniExcelLibs.Zip;

namespace MiniExcelLibs.OpenXml;

internal class ExcelOpenXmlStyles
{
    private static readonly string[] Ns = [Schemas.SpreadsheetmlXmlns, Schemas.SpreadsheetmlXmlStrictns];
    
    private readonly Dictionary<int, StyleRecord> _cellXfs = new();
    private readonly Dictionary<int, StyleRecord> _cellStyleXfs = new();
    private readonly Dictionary<int, NumberFormatString> _customFormats = new();

    public ExcelOpenXmlStyles(ExcelOpenXmlZip zip)
    {
        using var reader = zip.GetXmlReader("xl/styles.xml");
        
        if (!XmlReaderHelper.IsStartElement(reader, "styleSheet", Ns))
            return;
        if (!XmlReaderHelper.ReadFirstContent(reader))
            return;
                
        while (!reader.EOF)
        {
            if (XmlReaderHelper.IsStartElement(reader, "cellXfs", Ns))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                    continue;

                var index = 0;
                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "xf", Ns))
                    {
                        int.TryParse(reader.GetAttribute("xfId"), out var xfId);
                        int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);
                        _cellXfs.Add(index, new StyleRecord() { XfId = xfId, NumFmtId = numFmtId });
                        reader.Skip();
                        index++;
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                        break;
                }
            }
            else if (XmlReaderHelper.IsStartElement(reader, "cellStyleXfs", Ns))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                    continue;

                var index = 0;
                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "xf", Ns))
                    {
                        int.TryParse(reader.GetAttribute("xfId"), out var xfId);
                        int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);

                        _cellStyleXfs.Add(index, new StyleRecord() { XfId = xfId, NumFmtId = numFmtId });
                        reader.Skip();
                        index++;
                    }
                    else if (!XmlReaderHelper.SkipContent(reader))
                    {
                        break;
                    }
                }
            }
            else if (XmlReaderHelper.IsStartElement(reader, "numFmts", Ns))
            {
                if (!XmlReaderHelper.ReadFirstContent(reader))
                    continue;

                while (!reader.EOF)
                {
                    if (XmlReaderHelper.IsStartElement(reader, "numFmt", Ns))
                    {
                        int.TryParse(reader.GetAttribute("numFmtId"), out var numFmtId);
                        var formatCode = reader.GetAttribute("formatCode");
                                
                        //TODO: determine the type according to the format
                        var type = typeof(string);
                        if (DateTimeHelper.IsDateTimeFormat(formatCode))
                        {
                            type = typeof(DateTime?);
                        }

                        if (!_customFormats.ContainsKey(numFmtId))
                            _customFormats.Add(numFmtId, new NumberFormatString(formatCode, type));
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
                break;
            }
        }
    }

    public NumberFormatString? GetStyleFormat(int index)
    {
        if (!_cellXfs.TryGetValue(index, out var styleRecord)) 
            return null;
            
        if (Formats.TryGetValue(styleRecord.NumFmtId, out var numberFormat))
            return numberFormat;

        if (_customFormats.TryGetValue(styleRecord.NumFmtId, out var customNumberFormat))
            return customNumberFormat;
            
        return null;
    }

    public object? ConvertValueByStyleFormat(int index, object? value)
    {
        var sf = GetStyleFormat(index);
        
        if (sf?.Type is null)
            return value;

        if (sf.Type == typeof(DateTime?) && double.TryParse(value?.ToString(), out var s))
            return DateTimeHelper.IsValidOADateTime(s) ? DateTime.FromOADate(s) : value;

        if (sf.Type == typeof(TimeSpan?) && double.TryParse(value?.ToString(), out var number))
            return TimeSpan.FromDays(number);

        return value;
    }

    private static Dictionary<int, NumberFormatString> Formats { get; } = new()
    {
        { 0, new NumberFormatString("General", typeof(string)) },
        { 1, new NumberFormatString("0", typeof(decimal?)) },
        { 2, new NumberFormatString("0.00", typeof(decimal?)) },
        { 3, new NumberFormatString("#,##0", typeof(decimal?)) },
        { 4, new NumberFormatString("#,##0.00", typeof(decimal?)) },
        { 5, new NumberFormatString("\"$\"#,##0_);(\"$\"#,##0)", typeof(decimal?)) },
        { 6, new NumberFormatString("\"$\"#,##0_);[Red](\"$\"#,##0)", typeof(decimal?)) },
        { 7, new NumberFormatString("\"$\"#,##0.00_);(\"$\"#,##0.00)", typeof(decimal?)) },
        { 8, new NumberFormatString("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)", typeof(string)) },
        { 9, new NumberFormatString("0%", typeof(decimal?)) },
        { 10, new NumberFormatString("0.00%", typeof(string)) },
        { 11, new NumberFormatString("0.00E+00", typeof(string)) },
        { 12, new NumberFormatString("# ?/?", typeof(string)) },
        { 13, new NumberFormatString("# ??/??", typeof(string)) },
        { 14, new NumberFormatString("d/m/yyyy", typeof(DateTime?)) },
        { 15, new NumberFormatString("d-mmm-yy", typeof(DateTime?)) },
        { 16, new NumberFormatString("d-mmm", typeof(DateTime?)) },
        { 17, new NumberFormatString("mmm-yy", typeof(TimeSpan)) },
        { 18, new NumberFormatString("h:mm AM/PM", typeof(TimeSpan?)) },
        { 19, new NumberFormatString("h:mm:ss AM/PM", typeof(TimeSpan?)) },
        { 20, new NumberFormatString("h:mm", typeof(TimeSpan?)) },
        { 21, new NumberFormatString("h:mm:ss", typeof(TimeSpan?)) },
        { 22, new NumberFormatString("m/d/yy h:mm", typeof(DateTime?)) },
        // 23..36 international/unused
        { 37, new NumberFormatString("#,##0_);(#,##0)", typeof(string)) },
        { 38, new NumberFormatString("#,##0_);[Red](#,##0)", typeof(string)) },
        { 39, new NumberFormatString("#,##0.00_);(#,##0.00)", typeof(string)) },
        { 40, new NumberFormatString("#,##0.00_);[Red](#,##0.00)", typeof(string)) },
        { 41, new NumberFormatString("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)", typeof(string)) },
        { 42, new NumberFormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)", typeof(string)) },
        { 43, new NumberFormatString("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)", typeof(string)) },
        { 44, new NumberFormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)", typeof(string)) },
        { 45, new NumberFormatString("mm:ss", typeof(TimeSpan?)) },
        { 46, new NumberFormatString("[h]:mm:ss", typeof(TimeSpan?)) },
        { 47, new NumberFormatString("mm:ss.0", typeof(TimeSpan?)) },
        { 48, new NumberFormatString("##0.0E+0", typeof(string)) },
        { 49, new NumberFormatString("@", typeof(string)) },

        // issue 222
        { 58, new NumberFormatString("m/d",typeof(DateTime?)) },

        // custom format start with 176
    };
}

internal class NumberFormatString(string formatCode, Type type, bool needConvertToString = false)
{
    public string FormatCode { get; } = formatCode;
    public Type Type { get; set; } = type;
    public bool NeedConvertToString { get; } = needConvertToString;
}

internal class StyleRecord
{
    public int XfId { get; set; }
    public int NumFmtId { get; set; }
}