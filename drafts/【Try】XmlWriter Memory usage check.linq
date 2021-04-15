<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xml");
	Console.WriteLine(path);
	using(var stream = File.Create(path))
	using (XmlWriter writer = XmlWriter.Create(stream))
	{
		writer.WriteStartDocument();
		writer.WriteStartElement("root");
		for (int i = 0; i < 100000000; i++)
		{
			writer.WriteStartElement("element");
			writer.WriteString("content");
			writer.WriteEndElement();

			if (i % 10000000 == 0)
			{
				Console.WriteLine($"time.{i} memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
			}
		}
		writer.WriteEndElement();
		writer.WriteEndDocument();
		Console.WriteLine($"memory usage: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
	}
	//Console.WriteLine(builder);
	Console.WriteLine("ok");
}

// You can define other methods, fields, classes and namespaces here
