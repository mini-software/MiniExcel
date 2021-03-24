<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Xunit</Namespace>
  <Namespace>MiniExcelLibs.Utils</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

#load "xunit"

void Main()
{
	RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.
}

[Fact()]
public void TestReadFirstFromLargeFile()
{
	var path = @"D:\git\MiniExcel\samples\csv\TestLargeFile_1,000,000.csv";
	using (var stream = File.OpenRead(path))
	{
		var rows = MiniExcel.Query(stream, false).Take(2).ToList();
		Assert.Equal("Id", rows[0].A);
		Assert.Equal("Text", rows[0].B);
		Assert.Equal("0", rows[1].A);
		Assert.Equal("Hello World", rows[1].B);
	}
}

[Fact()]
public void TestReadHeader()
{
	var path = @"D:\git\MiniExcel\samples\csv\TestHeader.csv";
	using (var stream = File.OpenRead(path))
	{
		var rows = MiniExcel.Query(stream, true).ToList();
		Assert.Equal("A1", rows[0].Column1);
		Assert.Equal("B1", rows[0].Column2);
		Assert.Equal("A2", rows[1].Column1);
		Assert.Equal("B2", rows[1].Column2);
	}
}

// You can define other methods, fields, classes and namespaces here
public static class MiniExcel
{
	public static IEnumerable<dynamic> Query(this FileStream stream, bool useHeaderRow, IConfiguration configuration = null)
	{
		return CsvReader.Query(stream, useHeaderRow,(CsvConfiguration)configuration);
	}
}

public interface IConfiguration
{
}

public class CsvConfiguration : IConfiguration
{
	public char Seperator { get; set; }
	public Func<Stream,StreamReader> GetStreamReaderFunc { get; set; }
	private static readonly CsvConfiguration _defaultConfiguration = new CsvConfiguration()
	{
		Seperator = ',',
		GetStreamReaderFunc = (stream)=> new StreamReader(stream)
	};
	internal static CsvConfiguration GetDefaultConfiguration() => _defaultConfiguration;
}

public class CsvReader
{
	internal static IEnumerable<IDictionary<string, object>> Query(FileStream stream, bool useHeaderRow, CsvConfiguration configuration)
	{
		if (configuration == null)
			configuration = CsvConfiguration.GetDefaultConfiguration();
		using (var reader = configuration.GetStreamReaderFunc(stream))
		{
			char[] seperators = { configuration.Seperator };

			var row = string.Empty;
			string[] read;
			var firstRow = true;
			Dictionary<int, string> headRows = new Dictionary<int, string>();
			while ((row = reader.ReadLine()) != null)
			{
				read = row.Split(seperators, StringSplitOptions.None);

				//header
				if (useHeaderRow)
				{
					if (firstRow)
					{
						firstRow = false;
						for (int i = 0; i <= read.Length - 1; i++)
							headRows.Add(i, read[i]);
						continue;
					}

					var cell = Helpers.GetEmptyExpandoObject(headRows);
					for (int i = 0; i <= read.Length - 1; i++)
						cell[headRows[i]] = read[i];

					yield return cell;
					continue;
				}


				//body
				{
					var cell = Helpers.GetEmptyExpandoObject(read.Length - 1);
					for (int i = 0; i <= read.Length - 1; i++)
						cell[Helpers.GetAlphabetColumnName(i)] = read[i];
					yield return cell;
				}
			}
		}
	}
}

namespace MiniExcelLibs.Utils
{
	using System;
	using System.Collections;
	using System.Collections.Generic;
	using System.Dynamic;
	using System.Globalization;
	using System.Linq;
	using System.Reflection;
	using System.Text.RegularExpressions;

	internal static class Helpers
	{
		private static Dictionary<int, string> _IntMappingAlphabet = new Dictionary<int, string>();
		private static Dictionary<string, int> _AlphabetMappingInt = new Dictionary<string, int>();
		static Helpers()
		{
			for (int i = 0; i <= 255; i++)
			{
				_IntMappingAlphabet.Add(i, IntToLetters(i));
				_AlphabetMappingInt.Add(IntToLetters(i), i);
			}
		}

		public static string GetAlphabetColumnName(int ColumnIndex) => _IntMappingAlphabet[ColumnIndex];
		public static int GetColumnIndex(string columnName) => _AlphabetMappingInt[columnName];

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

		public static IDictionary<string, object> GetEmptyExpandoObject(int maxColumnIndex)
		{
			// TODO: strong type mapping can ignore this
			// TODO: it can recode better performance 
			var cell = (IDictionary<string, object>)new ExpandoObject();
			for (int i = 0; i <= maxColumnIndex; i++)
			{
				var key = GetAlphabetColumnName(i);
				if (!cell.ContainsKey(key))
					cell.Add(key, null);
			}
			return cell;
		}

		public static IDictionary<string, object> GetEmptyExpandoObject(Dictionary<int, string> hearrows)
		{
			// TODO: strong type mapping can ignore this
			// TODO: it can recode better performance 
			var cell = (IDictionary<string, object>)new ExpandoObject();
			foreach (var hr in hearrows)
				if (!cell.ContainsKey(hr.Value))
					cell.Add(hr.Value, null);
			return cell;
		}

		public static IEnumerable<PropertyInfo> GetPropertiesWithSetter(this Type type)
		{
			return type.GetProperties(BindingFlags.SetProperty |
					  BindingFlags.Public |
					  BindingFlags.Instance).Where(prop => prop.GetSetMethod() != null);
		}

		public static PropertyInfo[] GetSubtypeProperties(ICollection value)
		{
			var collectionType = value.GetType();

			Type gType;
			if (collectionType.IsGenericTypeDefinition || collectionType.IsGenericType)
				gType = collectionType.GetGenericArguments().Single();
			else if (collectionType.IsArray)
				gType = collectionType.GetElementType();
			else
				throw new NotImplementedException($"{collectionType.Name} type not implemented,please issue for me, https://github.com/shps951023/MiniExcel/issues");
			if (typeof(IDictionary).IsAssignableFrom(gType))
				throw new NotImplementedException($"{gType.Name} type not implemented,please issue for me, https://github.com/shps951023/MiniExcel/issues");
			var props = gType.GetProperties(BindingFlags.Public | BindingFlags.Instance);
			if (props.Length == 0)
				throw new InvalidOperationException($"Properties count is 0");
			return props;
		}

		private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");
		public static string ConvertEscapeChars(string input)
		{
			return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
		}

	}

}


