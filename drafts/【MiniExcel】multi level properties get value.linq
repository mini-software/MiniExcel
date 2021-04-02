<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Newtonsoft.Json.Linq</Namespace>
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
	var users = new List<dynamic>() { new { ID = 1, Name = "Jack" }, new { ID = 2, Name = "Lisa" } };
	var input = new
	{
		Title = "Demo Basic",
		Users = users,
		level1 = new
		{
			Users = users,
			level2 = new
			{
				Users = users,
				level3 = new
				{
					Users = users,
					value = "HelloWorld"
				}
			}
		}
	};

	Console.WriteLine(ReflectionHelper.GetValueByMultiPropertyName(input, "WrongName")); // null
	Console.WriteLine(ReflectionHelper.GetValueByMultiPropertyName(input, "Title")); //Demo Basic
	Console.WriteLine(ReflectionHelper.GetValueByMultiPropertyName(input, "level1.Users")); // OK
	Console.WriteLine(ReflectionHelper.GetValueByMultiPropertyName(input, "level1.level2")); // OK
	Console.WriteLine(ReflectionHelper.GetValueByMultiPropertyName(input, "level1.level2.level3.value")); // OK
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

