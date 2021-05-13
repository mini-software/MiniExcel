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

//[c# - Adding a prefix to an xml node - Stack Overflow](https://stackoverflow.com/questions/5157448/adding-a-prefix-to-an-xml-node)
void Main()
{
	{
		var xml = @"<Folio>
<Node1>Value1</Node1>
<Node2>Value2</Node2>
<Node3>Value3</Node3>
</Folio>";

		var doc = new XmlDocument();
		doc.LoadXml(xml);
		SetPrefix("x", doc.ChildNodes[0]);
		Console.WriteLine(doc.ChildNodes);
	}
	{
		var xml = @"<Folio xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<Node1>Value1</Node1>
<Node2>Value2</Node2>
<Node3>Value3</Node3>
</Folio>";

		var doc = new XmlDocument();
		doc.LoadXml(xml);
		SetPrefix("x", doc.ChildNodes[0]);
		Console.WriteLine(doc.ChildNodes);
	}
}

// You can define other methods, fields, classes and namespaces here
public static void SetPrefix(string prefix, XmlNode node)
{
	node.Prefix = prefix;
	foreach (XmlNode n in node.ChildNodes)
		SetPrefix(prefix, n);
}