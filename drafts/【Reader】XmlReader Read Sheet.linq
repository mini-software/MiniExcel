<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>System.Globalization</Namespace>
  <Namespace>System.Dynamic</Namespace>
  <Namespace>System.IO.Compression</Namespace>
</Query>

void Main()
{
	var path = @"D:\git\MiniExcel\samples\xlsx\TestDimensionCenterEmptyRows\TestDimensionCenterEmptyRows.xlsx";
	//using (var stream = File.OpenRead(path))
	//{
	//	var list = stream.Query();
	//	foreach (var e in list)
	//	{
	//		Console.WriteLine(e);
	//	}
	//}

	using (var stream = File.OpenRead(path))
	{
		Console.WriteLine(stream.Query());
	}

	using (var stream = File.OpenRead(path))
	{
		Console.WriteLine(stream.Query(true));
	}

	//Console.WriteLine("======");
	//using(var stream = File.OpenRead(path))
	//	Console.WriteLine(stream.GetValues(false).ToList());
}

public static class ExcelOpenXmlSheetReaderExtension
{
	public static IEnumerable<object> Query(this Stream stream, bool UseHeaderRow = false)
	{
		return new ExcelOpenXmlSheetReader().QueryImpl(stream, UseHeaderRow);
	}
}

internal class ExcelOpenXmlSheetReader
{
	internal readonly static XName T;
	static ExcelOpenXmlSheetReader()
	{
		T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
	}
	internal Dictionary<int, string> GetSharedStrings(ZipArchiveEntry sharedStringsEntry)
	{
		var xl = XElement.Load(sharedStringsEntry.Open());
		var ts = xl.Descendants(T).Select((s, i) => new { i, v = s.Value?.ToString() })
			 .ToDictionary(s => s.i, s => s.v)
		;
		return ts;
	}

	private static Dictionary<int, string> _SharedStrings;

	internal IEnumerable<object> QueryImpl(Stream stream, bool UseHeaderRow = false)
	{
		using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, UTF8Encoding.UTF8))
		{
			var e = archive.Entries.SingleOrDefault(w => w.FullName == "xl/sharedStrings.xml");
			_SharedStrings = GetSharedStrings(e);

			var firstSheetEntry = archive.Entries.First(w => w.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase));
			using (var firstSheetEntryStream = firstSheetEntry.Open())
			{
				using (XmlReader reader = XmlReader.Create(firstSheetEntryStream, XmlSettings))
				{
					var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
					if (!reader.IsStartElement("worksheet", ns))
						yield break;


					if (!XmlReaderHelper.ReadFirstContent(reader))
						yield break;

					var maxRowIndex = -1;
					var maxColumnIndex = -1;
					while (!reader.EOF)
					{
						//TODO: will dimension after sheetData?
						//this method logic depends on dimension to get maxcolumnIndex, if without dimension then it need to foreach all rows first time to get maxColumn and maxRowColumn
						if (reader.IsStartElement("dimension", ns))
						{
							var @ref = reader.GetAttribute("ref");
							if (string.IsNullOrEmpty(@ref))
								throw new InvalidOperationException("Without sheet dimension data");
							var rs = @ref.Split(":");
							if (ReferenceHelper.ParseReference(rs[1], out int cIndex, out int rIndex))
							{
								maxColumnIndex = cIndex - 1;
								maxRowIndex = rIndex - 1;
							}
							else
								throw new InvalidOperationException("Invaild sheet dimension start data");
						}
						if (reader.IsStartElement("sheetData", ns))
						{
							if (!XmlReaderHelper.ReadFirstContent(reader))
							{
								continue;
							}

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
										var cell = UseHeaderRow ? Helpers.GetEmptyExpandoObject(headRows):Helpers.GetEmptyExpandoObject(maxColumnIndex);
										var columnIndex = 0;
										while (!reader.EOF)
										{
											if (reader.IsStartElement("c", ns))
											{
												var cellValue = ReadCell(reader, columnIndex, out var _columnIndex);
												columnIndex = _columnIndex;

												//if not using First Head then using 1,2,3 as index
												if (UseHeaderRow)
												{
													if (rowIndex == 0)
														headRows.Add(columnIndex, cellValue.ToString());
													else
														cell[headRows[columnIndex]]= cellValue;
												}
												else
													cell[columnIndex.ToString()] = cellValue;
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
	}

	private object ReadCell(XmlReader reader, int nextColumnIndex, out int columnIndex)
	{
		var aT = reader.GetAttribute("t");
		var aR = reader.GetAttribute("r");

		//TODO:need to check only need nextColumnIndex or columnIndex
		if (ReferenceHelper.ParseReference(aR, out int referenceColumn, out _))
			columnIndex = referenceColumn - 1; // ParseReference is 1-based
		else
			columnIndex = nextColumnIndex;

		if (!XmlReaderHelper.ReadFirstContent(reader))
			return null;


		object value = null;
		while (!reader.EOF)
		{
			if (reader.IsStartElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				string rawValue = reader.ReadElementContentAsString();
				if (!string.IsNullOrEmpty(rawValue))
					ConvertCellValue(rawValue, aT, out value);
			}
			else if (reader.IsStartElement("is", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				string rawValue = StringHelper.ReadStringItem(reader);
				if (!string.IsNullOrEmpty(rawValue))
					ConvertCellValue(rawValue, aT, out value);
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}
		return value;
	}

	private void ConvertCellValue(string rawValue, string aT, out object value)
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


internal static class Helpers
{
	private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");

	public static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex)
	{
		// TODO: strong type mapping can ignore this
		// TODO: it can recode better performance 
		var cell = (IDictionary<string, object>)new ExpandoObject();
		for (int i = 0; i <= maxColumnIndex; i++)
			cell.Add(i.ToString(), null);
		return cell;
	}

	public static IDictionary<string, object> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
	{
		// TODO: strong type mapping can ignore this
		// TODO: it can recode better performance 
		var cell = (IDictionary<string, object>)new ExpandoObject();
		foreach (var hr in hearrows)
			cell.Add(hr.Value, null);
		return cell;
	}


	public static string ConvertEscapeChars(string input)
	{
		return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
	}

	/// <summary>
	/// Convert a double from Excel to an OA DateTime double. 
	/// The returned value is normalized to the '1900' date mode and adjusted for the 1900 leap year bug.
	/// </summary>
	public static double AdjustOADateTime(double value, bool date1904)
	{
		if (!date1904)
		{
			// Workaround for 1900 leap year bug in Excel
			if (value >= 0.0 && value < 60.0)
				return value + 1;
		}
		else
		{
			return value + 1462.0;
		}

		return value;
	}

	public static bool IsValidOADateTime(double value)
	{
		return value > DateTimeHelper.OADateMinAsDouble && value < DateTimeHelper.OADateMaxAsDouble;
	}

	public static object ConvertFromOATime(double value, bool date1904)
	{
		var dateValue = AdjustOADateTime(value, date1904);
		if (IsValidOADateTime(dateValue))
			return DateTimeHelper.FromOADate(dateValue);
		return value;
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

internal static class StringHelper
{
	public static string ReadStringItem(XmlReader reader)
	{
		string result = string.Empty;
		if (!XmlReaderHelper.ReadFirstContent(reader))
		{
			return result;
		}

		while (!reader.EOF)
		{
			if (reader.IsStartElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				// There are multiple <t> in a <si>. Concatenate <t> within an <si>.
				result += reader.ReadElementContentAsString();
			}
			else if (reader.IsStartElement("r", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				result += ReadRichTextRun(reader);
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}

		return result;
	}

	private static string ReadRichTextRun(XmlReader reader)
	{
		string result = string.Empty;
		if (!XmlReaderHelper.ReadFirstContent(reader))
		{
			return result;
		}

		while (!reader.EOF)
		{
			if (reader.IsStartElement("t", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
			{
				result += reader.ReadElementContentAsString();
			}
			else if (!XmlReaderHelper.SkipContent(reader))
			{
				break;
			}
		}

		return result;
	}
}

internal static class ReferenceHelper
{
	/// <summary>
	/// Logic for the Excel dimensions. Ex: A15
	/// </summary>
	/// <param name="value">The value.</param>
	/// <param name="column">The column, 1-based.</param>
	/// <param name="row">The row, 1-based.</param>
	public static bool ParseReference(string value, out int column, out int row)
	{
		column = 0;
		var position = 0;
		const int offset = 'A' - 1;

		if (value != null)
		{
			while (position < value.Length)
			{
				var c = value[position];
				if (c >= 'A' && c <= 'Z')
				{
					position++;
					column *= 26;
					column += c - offset;
					continue;
				}

				if (char.IsDigit(c))
					break;

				position = 0;
				break;
			}
		}

		if (position == 0)
		{
			column = 0;
			row = 0;
			return false;
		}

		if (!int.TryParse(value.Substring(position), NumberStyles.None, CultureInfo.InvariantCulture, out row))
		{
			return false;
		}

		return true;
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



