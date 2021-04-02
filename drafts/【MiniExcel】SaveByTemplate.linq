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
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var input = new Demo
	{
		Users = new User[] {
			new User{ID=Guid.NewGuid(),Name="Jack",Age=25,InDate=new DateTime(2021,3,1),VIP=true,Points=new Decimal(1234.55)},
			new User{ID=Guid.NewGuid(),Name="Lisa",Age=44,InDate=new DateTime(2021,2,14),VIP=false,Points=new Decimal(5741.201)},
		}
	};
	SaveByTemplate(input);
}

void SaveByTemplate(object input)
{
	var sheetXml = File.ReadAllText(@"D:\git\MiniExcel\samples\xlsx\TestBasicTemplate\xl\worksheets\sheet1.xml");
	var sharedXml = File.ReadAllText(@"D:\git\MiniExcel\samples\xlsx\TestBasicTemplate\xl\sharedStrings.xml");

	var shareds = new Dictionary<int, string>();
	{
		var xl = XElement.Parse(sharedXml);
		var ts = xl.Descendants(ExcelOpenXmlXName.T).Select((s, i) => new { i, v = s.Value?.ToString() })
			  .ToDictionary(s => s.i, s => s.v)
		;//TODO:need recode
		shareds = ts;
	}
	
	{
		// need to load all into memory first
		var sheetEXml = XElement.Parse(sheetXml);
		
		// if there is no next index row then create one 
		
		// set cell value
	}
}

internal class ReflectionHelper
{
	public class ReflectionResult
	{
		public string Key { get; set; }
		public bool isIEnumerable { get; set; }
		public object Value { get; set; }
		public PropertyInfo Property { get; set; }
	}

	// You can define other methods, fields, classes and namespaces here
	public static ReflectionResult GetValueByMultiPropertyName(object value, string name)
	{
		var result = new ReflectionResult();
		if (value == null)
			return result;
		var names = name.Split('.');
		Type t = null;
		result.Value = value;
		result.isIEnumerable = false;
		result.Key = name;
		result.Property = null;

		foreach (var n in names)
		{
			t = result.Value.GetType();
			result.Property = t.GetProperty(n);
			if (result.Property == null)
			{
				result.Value = null;
				return result;
			}

			result.isIEnumerable = typeof(IEnumerable).IsAssignableFrom(result.Property.PropertyType) && result.Property.PropertyType != typeof(string);
			result.Value = result.Property.GetValue(result.Value);
			if (result.Value == null)
			{
				result.Value = null;
				return result;
			}
		}

		return result;
	}

	public static bool TryGetValueByMultiPropertyName(object input, string name, out object value, out bool isIEnumerable)
	{
		isIEnumerable = false;
		if (input == null)
		{
			value = null;
			return false;
		}

		var names = name.Split('.');
		PropertyInfo p = null;
		Type t = null;
		value = input;

		foreach (var n in names)
		{
			t = value.GetType();
			p = t.GetProperty(n);
			if (p == null)
			{
				value = null;
				return false;
			}
			isIEnumerable = typeof(IEnumerable).IsAssignableFrom(p.PropertyType);
			Console.WriteLine($"{n} {isIEnumerable}");

			value = p.GetValue(value);
			if (value == null)
			{
				value = null;
				return false;
			}
		}
		return true;
	}
}

public class Demo
{
	public User[] Users { get; set; }
}
public class User
{
	public Guid ID { get; set; }
	public string Name { get; set; }
	public int Age { get; set; }
	public DateTime InDate { get; set; }
	public bool VIP { get; set; }
	public decimal Points { get; set; }
}



// You can define other methods, fields, classes and namespaces here
internal static class ExcelOpenXmlXName
{
	internal readonly static XNamespace ExcelNamespace = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
	internal readonly static XNamespace ExcelRelationshipsNamepace = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
	internal readonly static XName Row;
	internal readonly static XName R;
	internal readonly static XName V;
	internal readonly static XName T;
	internal readonly static XName C;
	internal readonly static XName Dimension;
	internal readonly static XName Sheet;
	static ExcelOpenXmlXName()
	{
		Row = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "row";
		R = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "r";
		V = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "v";
		T = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "t";
		C = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "c";
		Dimension = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "dimension";
		Sheet = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main") + "sheet";
	}
}