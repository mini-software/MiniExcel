<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>Magicodes.IE.Excel</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>Magicodes.ExporterAndImporter.Core</Namespace>
  <Namespace>Magicodes.ExporterAndImporter.Excel</Namespace>
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
	{
		var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xlsx");
		var templatePath = @"D:\git\MiniExcel\samples\xlsx\TestTemplateBasicIEmumerableFill_Magicodes_IE.xlsx";
		var value = new
		{
			employees = Enumerable.Range(1, 1000).Select(s => new { name = "Jack", department = "HR" }).ToList()
		};

		//创建Excel导出对象
		IExportFileByTemplate exporter = new ExcelExporter();
		exporter.ExportByTemplate(path, value, templatePath).GetAwaiter().GetResult();
	}
}

/*
error :  
![image](https://user-images.githubusercontent.com/12729184/114646437-d47c8c80-9d0d-11eb-8d5f-a78c61a84536.png)
![image](https://user-images.githubusercontent.com/12729184/114646441-d6465000-9d0d-11eb-887a-940413c2fbc8.png)
![image](https://user-images.githubusercontent.com/12729184/114646466-e0684e80-9d0d-11eb-8d81-f56903cc20fb.png)


*/

// You can define other methods, fields, classes and namespaces here
