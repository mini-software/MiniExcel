<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
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

//// [c# - yield return statement inside a using() { } block Disposes before executing - Stack Overflow](https://stackoverflow.com/questions/1539114/yield-return-statement-inside-a-using-block-disposes-before-executing)
void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.csv");
	var csv =
@"1,2
3,4";
	File.WriteAllText(path, csv);

	Console.WriteLine(Test1(path).ToList()); //Ok
	//Console.WriteLine(Test2(path).ToList()); //Error
	//Console.WriteLine(Test4(path).ToList()); //Error
	Console.WriteLine(Test6(path).ToList()); //OK
}

// You can define other methods, fields, classes and namespaces here

IEnumerable<string[]> Test1(string path)
{
	using (var stream = File.OpenRead(path))
	using (var reader = new StreamReader(stream))
	{
		var row = string.Empty;
		while ((row = reader.ReadLine()) != null)
			yield return row.Split(new[] { ',' }, StringSplitOptions.None);
	}
}

IEnumerable<string[]> Test4(string path)
{
	using (var stream = File.OpenRead(path))
	using (var reader = new StreamReader(stream))
		return Test5(reader).AsEnumerable();
}


IEnumerable<string[]> Test5(StreamReader reader)
{
	var row = string.Empty;
	while ((row = reader.ReadLine()) != null)
		yield return row.Split(new[] { ',' }, StringSplitOptions.None);
}

IEnumerable<string[]> Test6(string path)
{
	using (var stream = File.OpenRead(path))
	using (var reader = new StreamReader(stream))
		foreach (var element in Test7(reader))
		{
			yield return element;
		}

}

IEnumerable<string[]> Test7(StreamReader reader)
{
	var row = string.Empty;
	while ((row = reader.ReadLine()) != null)
		yield return row.Split(new[] { ',' }, StringSplitOptions.None);
}

IEnumerable<string[]> Test2(string path)
{
	using (var stream = File.OpenRead(path))
		return Test3(stream);
}

IEnumerable<string[]> Test3(Stream stream)
{
	using (var reader = new StreamReader(stream))
	{
		var row = string.Empty;
		while ((row = reader.ReadLine()) != null)
			yield return row.Split(new[] { ',' }, StringSplitOptions.None);
	}
}