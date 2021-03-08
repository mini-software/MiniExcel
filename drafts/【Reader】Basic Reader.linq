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
  <Namespace>Xunit</Namespace>
</Query>

#load "xunit"

#region private::Tests

[Fact]
void Test_NullRow()
{
	using (var reader = new XlsxEasyRowsValueReader(null))
	{
		Action testCode = () => { var v = reader[0]; };
		var ex = Record.Exception(testCode);
		Console.WriteLine();

		Assert.NotNull(ex);
		Assert.IsType<InvalidOperationException>(ex);
	}
}

[Fact]
void Test_ReaderGetName()
{
	using (var reader = new XlsxEasyRowsValueReader(null))
	{
		Assert.Equal("A", reader.GetName(0));
		Assert.Equal("B", reader.GetName(1));
		Assert.Equal("BD", reader.GetName(55));
	}
}

[Fact]
void Test_ReaderGetOriginal()
{
	using (var reader = new XlsxEasyRowsValueReader(null))
	{
		Assert.Equal(0, reader.GetOrdinal("A"));
		//Assert.Equal("B", reader.GetName(1));
		//Assert.Equal("BD", reader.GetName(55));
	}
}

#endregion

void Main()
{
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.

	Console.WriteLine("========");
	using (var reader = new XlsxEasyRowsValueReader(null))
	{
		var st = reader.GetSchemaTable();
		Console.WriteLine(st);
		//todo
		var dt = new DataTable();
		dt.Load(reader);
	}

	Console.WriteLine("========");
	using (var reader = new XlsxEasyRowsValueReader(null))
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

	Console.WriteLine("========");
	using (var reader = new XlsxEasyRowsValueReader(null))
	{
		for (int i = 0; i < reader.FieldCount; i++)
		{
			var name = reader.GetName(i);
			Console.Write($" | name : {name}");
			//var ordinal = reader.GetOrdinal(name);
			//Console.Write($" | ordinal : {ordinal}");
			Console.WriteLine();
		}
	}
}



public class XlsxEasyRowsValueReader : IDataReader
{
	private static Dictionary<int, Dictionary<int, object>> _Rows { get; set; } = new Dictionary<int, Dictionary<int, object>>() {
		{0,new Dictionary<int,object>(){{0,0},{3,3}}},
		{3,new Dictionary<int,object>(){{2,2}}},
	};
	
	public XlsxEasyRowsValueReader(string filePath)
	{
		
	}

	public int RowCount { get; set; } = 4;
	public int FieldCount { get; set; } = 4;
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

	public int GetValues(object[] values) {
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
		return v==null ? typeof(string) : v.GetType();
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
