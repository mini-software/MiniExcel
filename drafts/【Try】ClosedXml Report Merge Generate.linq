<Query Kind="Program">
  <NuGetReference>ClosedXML.Report</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>ClosedXML.Report</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
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
	const string outputFile = @".\Output\report.xlsx";
	var template = new XLTemplate(@"C:\Users\Wei\Downloads\ClosedXmlMergeTemplate.xlsx");

	var value = new
	{
		project = new[] {
			new {name = "項目1",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
			new {name = "項目2",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
			new {name = "項目3",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
			new {name = "項目4",content="[]內容1,[]內容2,[]內容3,[]內容4,[]內容5"},
		}
	};
	template.AddVariable(value);
	template.Generate();
	template.SaveAs(outputFile);

	//Show report
	Process.Start(new ProcessStartInfo(outputFile) { UseShellExecute = true });
}

// You can define other methods, fields, classes and namespaces here
