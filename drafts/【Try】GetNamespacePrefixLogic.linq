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
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	GetNamespaceOfPrefix();
	GetPrefixOfNamespace();
}

// You can define other methods, fields, classes and namespaces here
public static void GetNamespaceOfPrefix()
{
	Console.WriteLine("GetNamespaceOfPrefix");
	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book xmlns:bk='urn:samples' bk:ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"</book>");

	XmlNode root = doc.FirstChild;

	//Create a new attribute.
	string ns = root.GetNamespaceOfPrefix("bk");
	XmlNode attr = doc.CreateNode(XmlNodeType.Attribute, "genre", ns);
	attr.Value = "novel";

	//Add the attribute to the document.
	root.Attributes.SetNamedItem(attr);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}

public static void GetPrefixOfNamespace()
{
	Console.WriteLine("GetPrefixOfNamespace");
	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book xmlns:bk='urn:samples' bk:ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"</book>");

	XmlNode root = doc.FirstChild;

	//Create a new node.
	string prefix = root.GetPrefixOfNamespace("urn:samples");
	XmlElement elem = doc.CreateElement(prefix, "style", "urn:samples");
	elem.InnerText = "hardcover";

	Console.WriteLine(elem);

	//Add the node to the document.
	root.AppendChild(elem);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}