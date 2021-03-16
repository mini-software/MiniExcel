<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>System.Globalization</Namespace>
  <RemoveNamespace>System.Collections</RemoveNamespace>
  <RemoveNamespace>System.Collections.Generic</RemoveNamespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Reflection</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	Stopwatch sw = new Stopwatch();
	sw.Start();
	Console.WriteLine("start memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB");

	int maxCellColumn = -1;
	int maxRowCount = -1; // number of rows with cell records

	var path = @"C:\Users\Wei\Downloads\Test1,000,000x10\xl\worksheets\sheet1.xml";
	using (var stream = File.OpenRead(path))
	using (var reader = XmlReader.Create(stream,XmlSettings))
	{
		while (reader.Read()) //3784ms
		{
			//<dimension ref="A1:J1000000"/>
			if( reader.IsStartElement("c") ) //4246ms
			{
				//var r = reader.GetAttribute("r"); //4829ms
				if(ReferenceHelper.ParseReference(reader.GetAttribute("r"),out var column,out var row)) //5600ms
				{
					column = column - 1;
					row = row - 1;
					maxRowCount = Math.Max(maxRowCount, row);
					maxCellColumn = Math.Max(maxCellColumn, column); //5701ms
				}
			}
			else if (reader.IsStartElement("dimension")) //6159ms > 5999ms
			{
				var @ref = reader.GetAttribute("ref");
				if (string.IsNullOrEmpty(@ref))
					throw new InvalidOperationException("Without sheet dimension data");
				var rs = @ref.Split(':');
				if (ReferenceHelper.ParseReference(rs[1], out int cIndex, out int rIndex))
				{
					maxRowCount = cIndex - 1;
					maxCellColumn = rIndex - 1;
					break;
				}
				else
					throw new InvalidOperationException("Invaild sheet dimension start data");
			}
		}
		Console.WriteLine($"maxRowCount : {maxRowCount} , maxCellColumn : {maxCellColumn}");
		
		Console.WriteLine("end memory usage: " + System.Diagnostics.Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024) + $"MB & run time : {sw.ElapsedMilliseconds}ms");
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

// You can define other methods, fields, classes and namespaces here
private static readonly XmlReaderSettings XmlSettings = new XmlReaderSettings
{
	IgnoreComments = true,
	IgnoreWhitespace = true,
	XmlResolver = null,
};
