<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.ComponentModel</Namespace>
  <Namespace>System.Dynamic</Namespace>
  <Namespace>System.Globalization</Namespace>
  <Namespace>Xunit</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Diagnostics</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

#load "xunit"

public void Main()
{
	//RunTests();  // Call RunTests() or press Alt+Shift+T to initiate testing.

}

[Fact]
public void CustomAttributeWihoutVaildPropertiesTest()
{
	Assert.Throws<System.InvalidOperationException>(()=>Helpers.GetExcelCustomPropertyInfos(typeof(CustomAttributesWihoutVaildPropertiesTestPoco)));
}

[Fact]
public void CustomAttributesTest()
{
	var props = Helpers.GetExcelCustomPropertyInfos(typeof(CustomAttributesTestPoco));
	Assert.Equal(new[] {"Column1","Column2","Test4"},props.Select(s=>s.ExcelColumnName));
}

internal class Helpers
{
	internal class ExcelCustomPropertyInfo
	{
		public int? ExcelColumnIndex { get; set; }
		public string ExcelColumnName { get; set; }
		public PropertyInfo Property { get; set; }
		public Type ExcludeNullableType { get; set; }
		public bool Nullable { get; internal set; }
	}
	internal static List<ExcelCustomPropertyInfo> GetExcelCustomPropertyInfos(Type type)
	{
		var props = type.GetProperties(BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance)
			.Where(prop => prop.GetSetMethod() != null
				&&
				!(prop.GetCustomAttribute<ExcelIgnoreAttribute>()?.ExcelIgnore == true)
			) /*ignore without set*/
			// solve : https://github.com/shps951023/MiniExcel/issues/138
			.Select(p =>
			{
				var gt = Nullable.GetUnderlyingType(p.PropertyType);
				return new ExcelCustomPropertyInfo
				{
					Property = p,
					ExcludeNullableType = gt ?? p.PropertyType,
					Nullable = gt != null ? true : false
				};
			})
			.ToList();
		if (props.Count == 0)
			throw new InvalidOperationException($"{type.Name} un-ignore properties count can't be 0");

		foreach (var cp in props)
		{
			cp.ExcelColumnName = cp.Property.GetCustomAttribute<ExcelColumnNameAttribute>()?.ExcelColumnName;
			if(cp.ExcelColumnName==null)
				cp.ExcelColumnName = cp.Property.Name;
		}
		return props;
	}
}

public class CustomAttributesWihoutVaildPropertiesTestPoco
{
	[ExcelIgnore]
	public string Test3 { get; set; }
	public string Test5 { get; }
	public string Test6 { get; private set; }
}

public class CustomAttributesTestPoco
{
	[ExcelColumnName("Column1")]
	public DateTime Test1 { get; set; }
	[ExcelColumnName("Column2")]
	public string Test2 { get; set; }
	[ExcelIgnore]
	public string Test3 { get; set; }
	public string Test4 { get; set; }
	public string Test5 { get; }
	public string Test6 { get; private set; } 
}

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelColumnNameAttribute : Attribute
{
	public string ExcelColumnName { get; set; }
	public ExcelColumnNameAttribute(string excelColumnName) => ExcelColumnName = excelColumnName;
}

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
public class ExcelIgnoreAttribute : Attribute
{
	public bool ExcelIgnore { get; set; }
	public ExcelIgnoreAttribute(bool excelIgnore = true) => ExcelIgnore = excelIgnore;
}

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

	public static PropertyInfo[] GetProperties(this Type type)
	{
		return type.GetProperties(
					 BindingFlags.Public |
					 BindingFlags.Instance);
	}

	internal class PropertyInfoAndNullableUnderLyingType
	{
		public PropertyInfo Property { get; set; }
		public Type ExcludeNullableType { get; set; }
		public bool Nullable { get; internal set; }
	}


	public static PropertyInfoAndNullableUnderLyingType[] GetPropertiesWithSetterAndExcludeNullableType(this Type type)
	{
		return type.GetProperties(BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance)
			.Where(prop => prop.GetSetMethod() != null)
			// solve : https://github.com/shps951023/MiniExcel/issues/138
			.Select(p =>
			{
				var gt = Nullable.GetUnderlyingType(p.PropertyType);
				return new PropertyInfoAndNullableUnderLyingType
				{
					Property = p,
					ExcludeNullableType = gt ?? p.PropertyType,
					Nullable = gt != null ? true : false
				};
			})
			.ToArray();
	}

	internal static bool IsDapperRows<T>()
	{
		return typeof(IDictionary<string, object>).IsAssignableFrom(typeof(T));
	}

	private static readonly Regex EscapeRegex = new Regex("_x([0-9A-F]{4,4})_");
	public static string ConvertEscapeChars(string input)
	{
		return EscapeRegex.Replace(input, m => ((char)uint.Parse(m.Groups[1].Value, NumberStyles.HexNumber)).ToString());
	}

}

