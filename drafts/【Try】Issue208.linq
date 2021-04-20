<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	Console.WriteLine(Helpers.GetAlphabetColumnName(255));	
	Console.WriteLine(Helpers.GetAlphabetColumnName(256));
	Console.WriteLine(Helpers.GetAlphabetColumnName(257));
	Console.WriteLine(Helpers.GetAlphabetColumnName(258));
	Console.WriteLine(Helpers.GetAlphabetColumnName(16383));
	//Console.WriteLine(Helpers.GetAlphabetColumnName(16384));
	var _IntMappingAlphabet = Helpers._IntMappingAlphabet;
	var _AlphabetMappingInt = Helpers._AlphabetMappingInt;
}

// You can define other methods, fields, classes and namespaces here
internal static class Helpers
{
	private const int GENERAL_COLUMN_INDEX = 255;
	private const int MAX_COLUMN_INDEX = 16383;

	internal static Dictionary<int, string> _IntMappingAlphabet;
	internal static Dictionary<string, int> _AlphabetMappingInt;
	static Helpers()
	{
		if (_IntMappingAlphabet == null && _AlphabetMappingInt == null)
		{
			_IntMappingAlphabet = new Dictionary<int, string>();
			_AlphabetMappingInt = new Dictionary<string, int>();
			for (int i = 0; i <= GENERAL_COLUMN_INDEX; i++)
			{
				_IntMappingAlphabet.Add(i, IntToLetters(i));
				_AlphabetMappingInt.Add(IntToLetters(i), i);
			}
		}
	}

	public static string GetAlphabetColumnName(int columnIndex)
	{
		if (columnIndex >= _IntMappingAlphabet.Count)
		{
			if (columnIndex > MAX_COLUMN_INDEX)
				throw new InvalidDataException($"ColumnIndex {columnIndex} over excel vaild max index.");
			for (int i = _IntMappingAlphabet.Count; i <= columnIndex; i++)
			{
				_IntMappingAlphabet.Add(i, IntToLetters(i));
				_AlphabetMappingInt.Add(IntToLetters(i), i);
			}
		}
		return _IntMappingAlphabet[columnIndex];
	}
	public static int GetColumnIndex(string columnName)
	{
		var columnIndex = _AlphabetMappingInt[columnName];

		return columnIndex;
	}

	internal static string IntToLetters(int value)
	{
		value = value + 1;
		string result = string.Empty;
		while (--value >= 0)
		{
			result = (char)('A' + value % 26) + result;
			value /= 26;
		}
		return result;
	}
}