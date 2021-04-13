<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

/*

[XmlWriter.WriteAttributeString Method (System.Xml) | Microsoft Docs](https://docs.microsoft.com/en-us/dotnet/api/system.xml.xmlwriter.writeattributestring?redirectedfrom=MSDN&view=net-5.0#System_Xml_XmlWriter_WriteAttributeString_System_String_System_String_System_String_System_String_)

XmlWriter prefix
*/

void Main()
{

	XmlWriter writer = null;

	writer = XmlWriter.Create("sampledata.xml");

	// Write the root element.
	writer.WriteStartElement("book");

	// Write the xmlns:bk="urn:book" namespace declaration.
	writer.WriteAttributeString("xmlns", "bk", null, "urn:book");
	
	writer.WriteAttributeString("xmlns", "bk", null, "urn:book");

	// Write the bk:ISBN="1-800-925" attribute.
	writer.WriteAttributeString("ISBN", "urn:book", "1-800-925");

	writer.WriteElementString("price", "19.95");

	// Write the close tag for the root element.
	writer.WriteEndElement();

	// Write the XML to file and close the writer.
	writer.Flush();
	writer.Close();
	
	Console.WriteLine(File.ReadAllText("sampledata.xml"));
}

// You can define other methods, fields, classes and namespaces here
