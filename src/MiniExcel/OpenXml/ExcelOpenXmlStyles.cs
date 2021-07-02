namespace MiniExcelLibs.OpenXml
{
    using MiniExcelLibs.Utils;
    using MiniExcelLibs.Zip;
    using System;
    using System.Collections.Generic;

    internal class ExcelOpenXmlStyles
    {
        private const string _ns = Config.SpreadsheetmlXmlns;
        private Dictionary<int, StyleRecord> _cellXfs = new Dictionary<int, StyleRecord>();
        private Dictionary<int, StyleRecord> _cellStyleXfs = new Dictionary<int, StyleRecord>();
        private Dictionary<int, NumberFormatString> _customFormats = new Dictionary<int, NumberFormatString>();

        public ExcelOpenXmlStyles(ExcelOpenXmlZip zip)
        {
            using (var Reader = zip.GetXmlReader(@"xl/styles.xml"))
            {
                if (!Reader.IsStartElement("styleSheet", _ns))
                    return;
                if (!XmlReaderHelper.ReadFirstContent(Reader))
                    return;
                while (!Reader.EOF)
                {
                    if (Reader.IsStartElement("cellXfs", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(Reader))
                            continue;

                        var index = 0;
                        while (!Reader.EOF)
                        {
                            if (Reader.IsStartElement("xf", _ns))
                            {
                                int.TryParse(Reader.GetAttribute("xfId"), out var xfId);
                                int.TryParse(Reader.GetAttribute("numFmtId"), out var numFmtId);
                                _cellXfs.Add(index, new StyleRecord() { XfId = xfId, NumFmtId = numFmtId });
                                Reader.Skip();
                                index++;
                            }
                            else if (!XmlReaderHelper.SkipContent(Reader))
                                break;
                        }
                    }
                    else if (Reader.IsStartElement("cellStyleXfs", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(Reader))
                            continue;

                        var index = 0;
                        while (!Reader.EOF)
                        {
                            if (Reader.IsStartElement("xf", _ns))
                            {
                                int.TryParse(Reader.GetAttribute("xfId"), out var xfId);
                                int.TryParse(Reader.GetAttribute("numFmtId"), out var numFmtId);

                                _cellStyleXfs.Add(index, new StyleRecord() { XfId = xfId, NumFmtId = numFmtId });
                                Reader.Skip();
                                index++;
                            }
                            else if (!XmlReaderHelper.SkipContent(Reader))
                                break;
                        }
                    }
                    else if (Reader.IsStartElement("numFmts", _ns))
                    {
                        if (!XmlReaderHelper.ReadFirstContent(Reader))
                            continue;

                        while (!Reader.EOF)
                        {
                            if (Reader.IsStartElement("numFmt", _ns))
                            {
                                int.TryParse(Reader.GetAttribute("numFmtId"), out var numFmtId);
                                var formatCode = Reader.GetAttribute("formatCode");


                                //TODO: determine the type according to the format
                                var type = typeof(string);
                                if(DateTimeHelper.IsDateTimeFormat(formatCode))
                                {
                                    type = typeof(DateTime?);
                                }

                                _customFormats.Add(numFmtId,new NumberFormatString(formatCode, type));
                                Reader.Skip();
                            }
                            else if (!XmlReaderHelper.SkipContent(Reader))
                            {
                                break;
                            }
                        }
                    }
                    else if (!XmlReaderHelper.SkipContent(Reader))
                    {
                        break;
                    }
                }
            }
        }

        public NumberFormatString GetStyleFormat(int index)
        {
            if (_cellXfs.TryGetValue(index, out var styleRecord))
            {
                if (Formats.TryGetValue(styleRecord.NumFmtId, out var numberFormat))
                {
                    return numberFormat;
                }
                else if (_customFormats.TryGetValue(styleRecord.NumFmtId, out var customNumberFormat))
                {
                    return customNumberFormat;
                }
                return null;
            }
            return null;
        }

        public object ConvertValueByStyleFormat(int index, object value)
        {
            var sf = this.GetStyleFormat(index);
            if (sf == null)
                return value;
            if (sf.Type == null)
                return value;

            if (sf.Type == typeof(DateTime?))
            {
                if (double.TryParse(value?.ToString(), out var s))
                {
                    //TODO: if format like yyyy, it need to convert to double, not return datetime value




                    return DateTimeHelper.FromOADate(s);
                }
            }
            else if (sf.Type == typeof(TimeSpan?))
            {
                if (double.TryParse(value?.ToString(), out var number))
                {
                    return TimeSpan.FromDays(number);
                }
            }

            return value;
        }

        private static Dictionary<int, NumberFormatString> Formats { get; } = new Dictionary<int, NumberFormatString>()
        {
            { 0, new NumberFormatString("General",typeof(string)) },
            { 1, new NumberFormatString("0",typeof(decimal?)) },
            { 2, new NumberFormatString("0.00",typeof(decimal?)) },
            { 3, new NumberFormatString("#,##0",typeof(decimal?)) },
            { 4, new NumberFormatString("#,##0.00",typeof(decimal?)) },
            { 5, new NumberFormatString("\"$\"#,##0_);(\"$\"#,##0)",typeof(decimal?)) },
            { 6, new NumberFormatString("\"$\"#,##0_);[Red](\"$\"#,##0)",typeof(decimal?)) },
            { 7, new NumberFormatString("\"$\"#,##0.00_);(\"$\"#,##0.00)",typeof(decimal?)) },
            { 8, new NumberFormatString("\"$\"#,##0.00_);[Red](\"$\"#,##0.00)",typeof(string)) },
            { 9, new NumberFormatString("0%",typeof(decimal?)) },
            { 10, new NumberFormatString("0.00%",typeof(string)) },
            { 11, new NumberFormatString("0.00E+00",typeof(string)) },
            { 12, new NumberFormatString("# ?/?",typeof(string)) },
            { 13, new NumberFormatString("# ??/??",typeof(string)) },
            { 14, new NumberFormatString("d/m/yyyy",typeof(DateTime?)) },
            { 15, new NumberFormatString("d-mmm-yy",typeof(DateTime?)) },
            { 16, new NumberFormatString("d-mmm",typeof(DateTime?)) },
            { 17, new NumberFormatString("mmm-yy",typeof(TimeSpan)) },
            { 18, new NumberFormatString("h:mm AM/PM",typeof(TimeSpan?)) },
            { 19, new NumberFormatString("h:mm:ss AM/PM",typeof(TimeSpan?)) },
            { 20, new NumberFormatString("h:mm",typeof(TimeSpan?)) },
            { 21, new NumberFormatString("h:mm:ss",typeof(TimeSpan?)) },
            { 22, new NumberFormatString("m/d/yy h:mm",typeof(DateTime?)) },
            // 23..36 international/unused
            { 37, new NumberFormatString("#,##0_);(#,##0)",typeof(string)) },
            { 38, new NumberFormatString("#,##0_);[Red](#,##0)",typeof(string)) },
            { 39, new NumberFormatString("#,##0.00_);(#,##0.00)",typeof(string)) },
            { 40, new NumberFormatString("#,##0.00_);[Red](#,##0.00)",typeof(string)) },
            { 41, new NumberFormatString("_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)",typeof(string)) },
            { 42, new NumberFormatString("_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)",typeof(string)) },
            { 43, new NumberFormatString("_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)",typeof(string)) },
            { 44, new NumberFormatString("_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)",typeof(string)) },
            { 45, new NumberFormatString("mm:ss",typeof(TimeSpan?)) },
            { 46, new NumberFormatString("[h]:mm:ss",typeof(TimeSpan?)) },
            { 47, new NumberFormatString("mm:ss.0",typeof(TimeSpan?)) },
            { 48, new NumberFormatString("##0.0E+0",typeof(string)) },
            { 49, new NumberFormatString("@",typeof(string)) },

            // issue 222
            { 58, new NumberFormatString("m/d",typeof(DateTime?)) },

            // custom format start with 176
        };
    }

    internal class NumberFormatString
    {
        public string FormatCode { get; }
        public Type Type { get; set; }
        public bool NeedConvertToString { get; }

        public NumberFormatString(string formatCode, Type type,bool needConvertToString=false)
        {
            FormatCode = formatCode;
            Type = type;
            NeedConvertToString = needConvertToString;
        }
    }

    internal class StyleRecord
    {
        public int XfId { get; set; }
        public int NumFmtId { get; set; }
    }
}