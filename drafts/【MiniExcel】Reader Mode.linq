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
  <Namespace>System.IO.Compression</Namespace>
</Query>

void Main()
{
	using (var reader = new XlsxEasyRowsValueReader(@"D:\git\MiniExcel\samples\xlsx\TestCenterEmptyRow\TestCenterEmptyRow.xlsx"))
	{
		while (reader.Read())
		{
			Console.WriteLine($"row : {reader.CurrentRowIndex}");
			Console.Write("Cells : | ");
			for (int i = 0; i < reader.FieldCount; i++)
			{
				var v = reader.GetValue(i);
				Console.Write(v == null ? null : v);
				Console.Write(" | ");
			}
			Console.WriteLine();
		}
	}
}

internal class Worksheet
{
	public int RowCount { get; set; }
	public int FieldCount { get; set; }
	public Dictionary<int, Dictionary<int, object>> Rows { get; set; }
}

// You can define other methods, fields, classes and namespaces here
public class MiniExcel
{
	internal static Worksheet ConvertAsSheet(string path)
	{
		using (FileStream stream = new FileStream(path, FileMode.Open))
		using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, false, UTF8Encoding.UTF8))
		{
			//sharedStrings must in memory cache
			Dictionary<int, string> GetSharedStrings()
			{
				var sharedStringsEntry = archive.Entries.SingleOrDefault(w => w.FullName == "xl/sharedStrings.xml");
				var xml = ConvertToString(sharedStringsEntry);
				var xl = XElement.Parse(xml);
				var ts = xl.Descendants(ExcelXName.T).Select((s, i) => new { i, v = s.Value?.ToString() })
					.ToDictionary(s => s.i, s => s.v)
				;
				return ts;
			}

			var sharedStrings = GetSharedStrings();

			//notice: for performance just read first one and no care the order
			var rowIndexMaximum = int.MinValue;
			var columnIndexMaximum = int.MinValue;



			var datarows = new Dictionary<int, Dictionary<int, object>>();
			var firstSheetEntry = archive.Entries.First(w => w.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase));
			{
				var xml = ConvertToString(firstSheetEntry);
				var xl = XElement.Parse(xml);

				foreach (var row in xl.Descendants(ExcelXName.Row))
				{
					//
					var datarow = new Dictionary<int, object>();
					{
						var r = row.Attribute("r")?.Value?.ToString();

						var rowIndex = int.MinValue;
						if (int.TryParse(r, out var _rowIndex))
							rowIndex = _rowIndex - 1; // The row attribute is 1 - based				
						rowIndexMaximum = Math.Max(rowIndexMaximum, rowIndex);

						datarows.Add(rowIndex, datarow);
					}

					foreach (var cell in row.Descendants(ExcelXName.C))
					{
						var t = cell.Attribute("t")?.Value?.ToString();
						var v = cell.Descendants(ExcelXName.V).SingleOrDefault()?.Value;
						if (t == "s")
						{
							if (!string.IsNullOrEmpty(v))
								v = sharedStrings[int.Parse(v)];
						}

						var r = cell.Attribute("r")?.Value?.ToString();
						{
							var cellIndex = GetColumnIndex(r) - 1;
							columnIndexMaximum = Math.Max(columnIndexMaximum, cellIndex);

							datarow.Add(cellIndex, v);
						}
					}
				}
			}

			return new Worksheet { FieldCount=columnIndexMaximum+1,RowCount=rowIndexMaximum+1,Rows=datarows};
		}
	}

	private static string ConvertToString(ZipArchiveEntry entry)
	{
		if (entry == null)
			return null;
		using (var eStream = entry.Open())
		using (var reader = new StreamReader(eStream))
			return reader.ReadToEnd();
	}

	/// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
	internal static void ConvertCellToXY(string cell,out int x,out int y)
	{
		x=GetColumnIndex(cell);
		y=GetCellNumber(cell);
	}

	internal static int GetColumnIndex(string cell)
	{
		const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		const int mode = 26;

		var x = 0;
		var cellLetter = GetCellLetter(cell);
		//AA=27,ZZ=702
		for (int i = 0; i < cellLetter.Length; i++)
			x = x * mode + keys.IndexOf(cellLetter[i]);

		return x;
	}

	internal static int GetCellNumber(string cell)
	{
		if (string.IsNullOrEmpty(cell))
			throw new Exception("cell is null or empty");
		string cellNumber = string.Empty;
		for (int i = 0; i < cell.Length; i++)
		{
			if (Char.IsDigit(cell[i]))
				cellNumber += cell[i];
		}
		return int.Parse(cellNumber);
	}

	internal static string GetCellLetter(string cell)
	{
		string GetCellLetter = string.Empty;
		for (int i = 0; i < cell.Length; i++)
		{
			if (Char.IsLetter(cell[i]))
				GetCellLetter += cell[i];
		}
		return GetCellLetter;
	}
}

internal static class ExcelXName
{
	internal readonly static XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	internal readonly static XNamespace ExcelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	internal readonly static XName Row;
	internal readonly static XName R;
	internal readonly static XName V;
	internal readonly static XName T;
	internal readonly static XName C;
	internal readonly static XName Dimension;
	static ExcelXName()
	{
		Row = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "row";
		R = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "r";
		V = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "v";
		T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
		C = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "c";
		Dimension = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "dimension";
	}
}

public class XlsxEasyRowsValueReader : IDataReader
{
	private static Worksheet[] _Sheets = new Worksheet[1];
	private Dictionary<int, Dictionary<int, object>> _Rows {get{return _Sheets[0].Rows;}}

	public XlsxEasyRowsValueReader(string filePath)
	{
 		_Sheets[0] = MiniExcel.ConvertAsSheet(filePath);
	}

	public int RowCount {get{return _Sheets[0].RowCount;}}
	public int FieldCount {get{return _Sheets[0].FieldCount;}}
	public int Depth { get; private set; }
	public int CurrentRowIndex { get { return Depth - 1; } }

	public object this[int i] => GetValue(i);
	public object this[string name] => GetValue(GetOrdinal(name));

	public bool Read()
	{
		if (Depth == RowCount)
			return false;
		Depth++;
		return true;
	}

	public string GetName(int i) => Helper.ConvertColumnName(i + 1);


	public int GetOrdinal(string name)
	{
		//TODO
		var keys = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray().ToList();
		var dic = keys.ToDictionary(s => s, s => keys.IndexOf(s));
		return dic[(name[0])];
	}

	public object GetValue(int i)
	{
		//if (CurrentRowIndex < 0)
		//	throw new InvalidOperationException("Invalid attempt to read when no data is present.");
		if (!_Rows.Keys.Contains(CurrentRowIndex))
			return null;
		if (_Rows[this.CurrentRowIndex].TryGetValue(i, out var v))
			return v;
		return null;
	}

	public int GetValues(object[] values)
	{
		return this.Depth;
	}

	//TODO: multiple sheets
	public bool NextResult() => false;

	public void Dispose() { }

	public void Close() { }

	public int RecordsAffected => throw new NotImplementedException();

	bool IDataReader.IsClosed => this.RowCount - 1 == this.Depth;

	public string GetString(int i) => (string)GetValue(i);

	public bool GetBoolean(int i) => (bool)GetValue(i);

	public byte GetByte(int i) => (byte)GetValue(i);

	public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length) => throw new NotImplementedException();

	public char GetChar(int i) => (char)GetValue(i);

	public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length) => throw new NotImplementedException();

	public IDataReader GetData(int i) => throw new NotImplementedException();

	public string GetDataTypeName(int i) => throw new NotImplementedException();

	public DateTime GetDateTime(int i) => (DateTime)GetValue(i);

	public decimal GetDecimal(int i) => (decimal)GetValue(i);

	public double GetDouble(int i) => (double)GetValue(i);

	public Type GetFieldType(int i)
	{
		var v = GetValue(i);
		return v == null ? typeof(string) : v.GetType();
	}

	public float GetFloat(int i) => (float)GetValue(i);

	public Guid GetGuid(int i) => (Guid)GetValue(i);

	public short GetInt16(int i) => (short)GetValue(i);

	public int GetInt32(int i) => (int)GetValue(i);

	public long GetInt64(int i) => (long)GetValue(i);

	public DataTable GetSchemaTable()
	{
		var dataTable = new DataTable("SchemaTable");
		dataTable.Locale = System.Globalization.CultureInfo.InvariantCulture;
		dataTable.Columns.Add("ColumnName", typeof(string));
		dataTable.Columns.Add("ColumnOrdinal", typeof(int));
		for (int i = 0; i < this.FieldCount; i++)
		{
			dataTable.Rows.Add(this.GetName(i), i);
		}
		DataColumnCollection columns = dataTable.Columns;
		foreach (DataColumn item in columns)
		{
			item.ReadOnly = true;
		}
		return dataTable;
	}

	public bool IsDBNull(int i) => GetValue(i) == null;

}

internal static class Helper
{
	public static TValue GetValueOrDefault<TKey, TValue>
	(this IDictionary<TKey, TValue> dictionary,
	 TKey key,
	 TValue defaultValue)
	{
		TValue value;
		return dictionary.TryGetValue(key, out value) ? value : defaultValue;
	}

	public static TValue GetValueOrDefault<TKey, TValue>
		(this IDictionary<TKey, TValue> dictionary,
		 TKey key,
		 Func<TValue> defaultValueProvider)
	{
		TValue value;
		return dictionary.TryGetValue(key, out value) ? value
			 : defaultValueProvider();
	}

	internal static string ConvertColumnName(int x)
	{
		int dividend = x;
		string columnName = String.Empty;
		int modulo;

		while (dividend > 0)
		{
			modulo = (dividend - 1) % 26;
			columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
			dividend = (int)((dividend - modulo) / 26);
		}
		return columnName;
	}
}