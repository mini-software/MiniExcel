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
  <RemoveNamespace>System.Text</RemoveNamespace>
  <RemoveNamespace>System.Text.RegularExpressions</RemoveNamespace>
  <RemoveNamespace>System.Threading</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.Linq</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	Clone();
	Console.WriteLine();
	appendchild();
	Console.WriteLine();
	prependchild();
	Console.WriteLine();
	insertafter();
}



public static void Clone()
{
	Console.WriteLine($"==== Clone  ====");

	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"<title2>Pride And Prejudice</title2>" +
				"<title3>Pride And Prejudice</title3>" +
				"</book>");

	XmlNode root = doc.DocumentElement;

	//Create a new node.
	var elem = doc.SelectSingleNode("/book/title");
	root.InsertAfter(elem.Clone(),elem);
	root.InsertAfter(elem.Clone(),elem);
	root.InsertAfter(elem.Clone(),elem);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}

public static void appendchild()
{
	Console.WriteLine($"==== appendchild ====");

	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"<title2>Pride And Prejudice</title2>" +
				"<title3>Pride And Prejudice</title3>" +
				"</book>");

	XmlNode root = doc.DocumentElement;

	//Create a new node.
	XmlElement elem = doc.CreateElement("price");
	elem.InnerText = "19.95";

	//Add the node to the document.
	root.AppendChild(elem);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}

// Adds the specified node to the "beginning" of the list of child nodes for this node.
public static void prependchild()
{
	Console.WriteLine("==== prependchild ====");
	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"</book>");

	XmlNode root = doc.DocumentElement;

	//Create a new node.
	XmlElement elem = doc.CreateElement("price");
	elem.InnerText = "19.95";

	//Add the node to the document.
	root.PrependChild(elem);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}

public static void insertafter()
{
	Console.WriteLine("==== insertafter ====");
	XmlDocument doc = new XmlDocument();
	doc.LoadXml("<book genre='novel' ISBN='1-861001-57-5'>" +
				"<title>Pride And Prejudice</title>" +
				"<title2>Pride And Prejudice</title2>" +
				"<title3>Pride And Prejudice</title3>" +
				"</book>");

	XmlNode root = doc.DocumentElement;

	//Create a new node.
	XmlElement elem = doc.CreateElement("price");
	elem.InnerText = "19.95";

	//Add the node to the document.
	root.InsertAfter(elem, root.FirstChild);

	Console.WriteLine("Display the modified XML...");
	doc.Save(Console.Out);
}


// You can define other methods, fields, classes and namespaces here
