<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>Xunit</Namespace>
  <Namespace>MiniExcelLibs.OpenXml</Namespace>
</Query>

#load "xunit"

void Main()
{
	var path = @"D:\git\MiniExcel\MiniExcel\samples\xlsx\TestDatetimeSpanFormat_ClosedXml.xlsx";
	using (var stream = File.OpenRead(path))
	using (var zip = new ExcelOpenXmlZip(stream))
	{
		var style = new ExcelOpenXmlStyles(zip);
		var value = style.ConvertValueByStyleFormat(1,"44274.8758969329");	
		Console.WriteLine(value);
	}
}

namespace MiniExcelLibs.OpenXml
{
	using System;
	using System.Collections.Generic;
	using System.Xml;
	internal class ExcelOpenXmlStyles
	{
		const string NsSpreadsheetMl = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

		private static Dictionary<int, StyleRecord> _cellXfs = new Dictionary<int, StyleRecord>();
		private static Dictionary<int, StyleRecord> _cellStyleXfs = new Dictionary<int, StyleRecord>();

		private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
		{
			IgnoreComments = true,
			IgnoreWhitespace = true,
			XmlResolver = null,
		};

		public ExcelOpenXmlStyles(ExcelOpenXmlZip zip)
		{
			using (var Reader = zip.GetXmlReader(@"xl/styles.xml"))
			{
				if (!Reader.IsStartElement("styleSheet", NsSpreadsheetMl))
					return;
				if (!XmlReaderHelper.ReadFirstContent(Reader))
					return;
				while (!Reader.EOF)
				{
					if (Reader.IsStartElement("cellXfs", NsSpreadsheetMl))
					{
						if (!XmlReaderHelper.ReadFirstContent(Reader))
							return;

						var index = 0;
						while (!Reader.EOF)
						{
							if (Reader.IsStartElement("xf", NsSpreadsheetMl))
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
					else if (Reader.IsStartElement("cellStyleXfs", NsSpreadsheetMl))
					{
						if (!XmlReaderHelper.ReadFirstContent(Reader))
							return;

						var index = 0;
						while (!Reader.EOF)
						{
							if (Reader.IsStartElement("xf", NsSpreadsheetMl))
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
				return null;
			}
			return null;
		}

		public object ConvertValueByStyleFormat(int index, object value)
		{
			var sf = this.GetStyleFormat(index);
			if (sf.Type == typeof(DateTime?))
			{
				if (double.TryParse(value?.ToString(), out var s))
				{
					return DateTimeHelper.FromOADate(s);
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
			{ 18, new NumberFormatString("h:mm AM/PM",typeof(TimeSpan)) },
			{ 19, new NumberFormatString("h:mm:ss AM/PM",typeof(TimeSpan)) },
			{ 20, new NumberFormatString("h:mm",typeof(TimeSpan)) },
			{ 21, new NumberFormatString("h:mm:ss",typeof(TimeSpan)) },
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
			{ 45, new NumberFormatString("mm:ss",typeof(TimeSpan)) },
			{ 46, new NumberFormatString("[h]:mm:ss",typeof(TimeSpan)) },
			{ 47, new NumberFormatString("mm:ss.0",typeof(TimeSpan)) },
			{ 48, new NumberFormatString("##0.0E+0",typeof(string)) },
			{ 49, new NumberFormatString("@",typeof(string)) },
		};
	}

	internal class NumberFormatString
	{
		public string FormatString { get; }
		public Type Type { get; set; }
		public NumberFormatString(string formatString, Type type)
		{
			FormatString = formatString;
			Type = type;
		}
	}

	internal class StyleRecord
	{
		public int XfId { get; set; }
		public int NumFmtId { get; set; }
	}
}

/// Copy & modified by ExcelDataReader ZipWorker
internal class ExcelOpenXmlZip : IDisposable
{
	private readonly Dictionary<string, ZipArchiveEntry> _entries;
	private bool _disposed;
	private Stream _zipStream;
	private ZipArchive _zipFile;
	private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
	{
		IgnoreComments = true,
		IgnoreWhitespace = true,
		XmlResolver = null,
	};
	public ExcelOpenXmlZip(Stream fileStream)
	{
		_zipStream = fileStream ?? throw new ArgumentNullException(nameof(fileStream));
		_zipFile = new ZipArchive(fileStream);
		_entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
		foreach (var entry in _zipFile.Entries)
		{
			_entries.Add(entry.FullName.Replace('\\', '/'), entry);
		}
	}

	private ZipArchiveEntry GetEntry(string path)
	{
		if (_entries.TryGetValue(path, out var entry))
			return entry;
		return null;
	}

	public XmlReader GetXmlReader(string path)
	{
		var entry = GetEntry(path);
		if (entry != null)
			return XmlReader.Create(entry.Open(), XmlSettings);
		return null;
	}

	~ExcelOpenXmlZip()
	{
		Dispose(false);
	}

	public void Dispose()
	{
		Dispose(true);

		GC.SuppressFinalize(this);
	}

	private void Dispose(bool disposing)
	{
		// Check to see if Dispose has already been called.
		if (!_disposed)
		{
			if (disposing)
			{
				if (_zipFile != null)
				{
					_zipFile.Dispose();
					_zipFile = null;
				}

				if (_zipStream != null)
				{
					_zipStream.Dispose();
					_zipStream = null;
				}
			}

			_disposed = true;
		}
	}
}


internal static class XmlReaderHelper
{
	public static bool ReadFirstContent(XmlReader xmlReader)
	{
		if (xmlReader.IsEmptyElement)
		{
			xmlReader.Read();
			return false;
		}

		xmlReader.MoveToContent();
		xmlReader.Read();
		return true;
	}

	public static bool SkipContent(XmlReader xmlReader)
	{
		if (xmlReader.NodeType == XmlNodeType.EndElement)
		{
			xmlReader.Read();
			return false;
		}

		xmlReader.Skip();
		return true;
	}
}

internal static class DateTimeHelper
{
	// All OA dates must be greater than (not >=) OADateMinAsDouble
	public const double OADateMinAsDouble = -657435.0;

	// All OA dates must be less than (not <=) OADateMaxAsDouble
	public const double OADateMaxAsDouble = 2958466.0;

	// From DateTime class to enable OADate in PCL
	// Number of 100ns ticks per time unit
	private const long TicksPerMillisecond = 10000;
	private const long TicksPerSecond = TicksPerMillisecond * 1000;
	private const long TicksPerMinute = TicksPerSecond * 60;
	private const long TicksPerHour = TicksPerMinute * 60;
	private const long TicksPerDay = TicksPerHour * 24;

	// Number of milliseconds per time unit
	private const int MillisPerSecond = 1000;
	private const int MillisPerMinute = MillisPerSecond * 60;
	private const int MillisPerHour = MillisPerMinute * 60;
	private const int MillisPerDay = MillisPerHour * 24;

	// Number of days in a non-leap year
	private const int DaysPerYear = 365;

	// Number of days in 4 years
	private const int DaysPer4Years = DaysPerYear * 4 + 1;

	// Number of days in 100 years
	private const int DaysPer100Years = DaysPer4Years * 25 - 1;

	// Number of days in 400 years
	private const int DaysPer400Years = DaysPer100Years * 4 + 1;

	// Number of days from 1/1/0001 to 12/30/1899
	private const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

	// Number of days from 1/1/0001 to 12/31/9999
	private const int DaysTo10000 = DaysPer400Years * 25 - 366;

	private const long MaxMillis = (long)DaysTo10000 * MillisPerDay;

	private const long DoubleDateOffset = DaysTo1899 * TicksPerDay;

	public static DateTime FromOADate(double d)
	{
		return new DateTime(DoubleDateToTicks(d), DateTimeKind.Unspecified);
	}

	// duplicated from DateTime
	internal static long DoubleDateToTicks(double value)
	{
		if (value >= OADateMaxAsDouble || value <= OADateMinAsDouble)
			throw new ArgumentException("Invalid OA Date");
		long millis = (long)(value * MillisPerDay + (value >= 0 ? 0.5 : -0.5));

		// The interesting thing here is when you have a value like 12.5 it all positive 12 days and 12 hours from 01/01/1899
		// However if you a value of -12.25 it is minus 12 days but still positive 6 hours, almost as though you meant -11.75 all negative
		// This line below fixes up the millis in the negative case
		if (millis < 0)
		{
			millis -= millis % MillisPerDay * 2;
		}

		millis += DoubleDateOffset / TicksPerMillisecond;

		if (millis < 0 || millis >= MaxMillis)
			throw new ArgumentException("OA Date out of range");
		return millis * TicksPerMillisecond;
	}
}