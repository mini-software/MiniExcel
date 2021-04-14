<Query Kind="Program">
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>MiniExcel</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcelLibs</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>MiniExcelLibs.OpenXml</Namespace>
  <RemoveNamespace>System.Data</RemoveNamespace>
  <RemoveNamespace>System.Linq.Expressions</RemoveNamespace>
  <RemoveNamespace>System.Transactions</RemoveNamespace>
  <RemoveNamespace>System.Xml.XPath</RemoveNamespace>
</Query>

void Main()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid().ToString()}.xml");
	Dictionary<string, object> Maps = new Dictionary<string, object>()
	{
		["title"] = "FooCompany",
		["date"] = DateTime.Now,
		["managers"] = new[] {
		new {name="Jack",department="HR"},
		new {name="Loan",department="IT"}
	},
		["employees"] = new[] {
		new {name="Wade",department="HR"},
		new {name="Felix",department="HR"},
		new {name="Eric",department="IT"},
		new {name="Keaton",department="IT"}
	}
	};
	using (var stream = File.Create(path))
	{
		ExcelOpenXmlTemplate.GenerateSheetXml(stream, xml, Maps);
		Console.WriteLine(File.ReadAllText(path));
	}
}

namespace MiniExcelLibs.OpenXml
{
	using System;
	using System.Collections;
	using System.Collections.Generic;
	using System.IO;
	using System.Linq;
	using System.Reflection;
	using System.Text;
	using System.Text.RegularExpressions;
	using System.Xml;
	using System.Xml.XPath;

	internal class ExcelOpenXmlTemplate
	{
		private const string _ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		private static readonly Regex _isExpressionRegex;
		static ExcelOpenXmlTemplate()
		{
			_isExpressionRegex = new Regex("(?<={{).*?(?=}})");
		}
		public static void GenerateSheetXml(Stream stream, string sheetXml, Dictionary<string, object> inputMaps, XmlWriterSettings xmlWriterSettings = null)
		{
			var doc = XDocument.Parse(sheetXml);
			

			XmlNamespaceManager ns = new XmlNamespaceManager(new NameTable());
			ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

			var worksheet = doc.XPathSelectElement("/x:worksheet", ns);

			XElement dimension = doc.XPathSelectElement("/x:worksheet/x:dimension", ns);
			XElement newDimension = null;
			if (dimension != null)
			{
				newDimension = new XElement(dimension);
				dimension.Remove();
			}


			var sheetData = doc.XPathSelectElement("/x:worksheet/x:sheetData", ns);
			var newSheetData = new XElement(sheetData);
			sheetData.RemoveAll();
			sheetData.Value = ("{{{{{{split}}}}}}");
			
		

			var contents = doc.ToString().Split(new string[] { "<sheetData>{{{{{{split}}}}}}</sheetData>" }, StringSplitOptions.None); 
			using (var writer = new StreamWriter(stream, Encoding.UTF8))
			{
				writer.Write(contents[0]);
				var rows = newSheetData.XPathSelectElements($"x:row", ns);

				//update dimension
				if (newDimension != null)
				{
					var maxRowIndexDiff = 0;
					foreach (var row in rows)
					{
						IEnumerable ienumerable = null;

						foreach (var c in row.XPathSelectElements($"x:c", ns))
						{
							var v = c.XPathSelectElement("x:v", ns);
							if (v?.Value == null)
								continue;

							var matchs = (_isExpressionRegex.Matches(v.ToString()).Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value));
							foreach (var item in matchs)
							{
								var keys = item.Split('.');
								var value = inputMaps[keys[0]];
								if (value is IEnumerable && !(value is string))
								{
									ienumerable = value as IEnumerable;
									break;
								}
							}
							if (ienumerable != null)
								break;
						}
						if (ienumerable != null)
						{
							var first = true;
							foreach (var element in ienumerable)
							{
								if (!first)
									maxRowIndexDiff++;
								first = false;
							}
						}
					}
					// e.g <dimension ref=\"A1:B6\" /> only need to update B6 to BMaxRowIndex
					var refAtt = newDimension.Attribute("ref");
					var @refs = refAtt.Value.Split(':');
					var letter = new String(refs[1].Where(Char.IsLetter).ToArray());
					var digit = int.Parse(new String(refs[1].Where(Char.IsDigit).ToArray()));

					refAtt.Value = $"{refs[0]}:{letter}{digit + maxRowIndexDiff}";
					writer.Write(CleanXml(newDimension.ToString()));
				}

				//render sheetData
//				writer.Write("<sheetData>");
//				int originRowIndex;
//				int rowIndexDiff = 0;
//				foreach (XmlElement row in rows)
//				{
//					var rowCotainIEnumerable = false;
//					IEnumerable ienumerable = null;
//					List<string> propKeys = null;
//					Type ienumerableGenricType = null;
//					string ienumerableKey = null;
//
//					//TODO: some xlsx without r
//					originRowIndex = int.Parse(row.GetAttribute("r"));
//
//					//TODO: need to remove namespace
//
//					// check if contains IEnumerble row
//					{
//						var cs = row.SelectNodes($"x:c", ns);
//						foreach (XmlElement c in cs)
//						{
//							var cr = c.GetAttribute("r");
//							var letter = new String(cr.Where(Char.IsLetter).ToArray());
//							c.SetAttribute("r", $"{letter}{{{{{{MiniExcel_RowIndex}}}}}}");
//
//							var v = c.SelectSingleNode("x:v", ns);
//							if (v?.InnerText == null)
//								continue;
//
//							var matchs = (_isExpressionRegex.Matches(v.InnerText).Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value));
//							foreach (var item in matchs)
//							{
//								var keys = item.Split('.');
//								var value = inputMaps[keys[0]];
//								if (value is IEnumerable && !(value is string))
//								{
//
//
//									if (propKeys == null)
//										propKeys = new List<string>();
//									propKeys.Add(keys[1]); //TODO: check if not contain 1 index
//														   //only check first one match IEnumerable, so only render one collection at same row
//									if (rowCotainIEnumerable == false)
//									{
//										ienumerableKey = keys[0];
//										// get ienumerable runtime type
//										foreach (var element in value as IEnumerable)
//										{
//											if (element != null)
//											{
//												ienumerableGenricType = element.GetType();
//												break;
//											}
//										}
//
//										ienumerable = value as IEnumerable;
//										rowCotainIEnumerable = true;
//									}
//								}
//								else
//								{
//									v.InnerText = v.InnerText.Replace($"{{{{{keys[0]}}}}}", value.ToString()); //TODO: auto check type and set value
//								}
//							}
//						}
//					}
//
//
//
//					var newRowIndex = originRowIndex + rowIndexDiff;
//					if (rowCotainIEnumerable && ienumerable != null)
//					{
//						var first = true;
//						foreach (var element in ienumerable)
//						{
//							var newRow = row.Clone() as XmlElement;
//							newRow.SetAttribute("r", newRowIndex.ToString());
//							newRow.InnerXml = row.InnerXml.Replace($"{{{{{{MiniExcel_RowIndex}}}}}}", newRowIndex.ToString());
//
//							foreach (var key in propKeys)
//							{
//								var prop = ienumerableGenricType.GetProperty(key);
//								newRow.InnerXml = newRow.InnerXml.Replace($"{{{{{ienumerableKey}.{key}}}}}", prop.GetValue(element).ToString());
//							}
//
//							// note: only first time need add diff ![image](https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png)
//							if (!first)
//								rowIndexDiff++;
//							first = false;
//
//							newRowIndex++;
//							writer.Write(CleanXml(newRow.OuterXml));
//						}
//					}
//					else
//					{
//						row.SetAttribute("r", newRowIndex.ToString());
//						row.InnerXml = row.InnerXml.Replace($"{{{{{{MiniExcel_RowIndex}}}}}}", newRowIndex.ToString());
//						writer.Write(CleanXml(row.OuterXml));
//					}
//
//				}
//				writer.Write("</sheetData>");
//				writer.Write(contents[1]);
			}
		}
		private static string CleanXml(string xml)
		{
			//TODO: need to optimize
			return xml.Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "").Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
		}
	}
}



const string xml = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
    xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:mc=""http://schemas.openxmlformats.org/markup-compatibility/2006"" mc:Ignorable=""x14ac xr xr2 xr3""
    xmlns:x14ac=""http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac""
    xmlns:xr=""http://schemas.microsoft.com/office/spreadsheetml/2014/revision""
    xmlns:xr2=""http://schemas.microsoft.com/office/spreadsheetml/2015/revision2""
    xmlns:xr3=""http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"" xr:uid=""{00000000-0001-0000-0000-000000000000}"">
    <dimension ref=""A1:B100"" />
    <sheetViews>
        <sheetView tabSelected=""1"" zoomScaleNormal=""100"" workbookViewId=""0"">
            <selection activeCell=""C8"" sqref=""C8"" />
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultColWidth=""11.5546875"" defaultRowHeight=""13.2"" x14ac:dyDescent=""0.25"" />
    <cols>
        <col min=""1"" max=""1"" width=""17.109375"" customWidth=""1"" />
        <col min=""2"" max=""2"" width=""21.77734375"" customWidth=""1"" />
    </cols>
    <sheetData>
        <row r=""1"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A1"" t=""str"">
                <v>Sheet1</v>
            </c>
        </row>
        <row r=""2"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A2"" t=""str"">
                <v>{{title}}-{{date}}</v>
            </c>
            <c r=""B2"" s=""2"" />
        </row>
        <row r=""3"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A3"" t=""str"">
                <v>Managers</v>
            </c>
        </row>
        <row r=""4"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A4"" t=""str"">
                <v>{{managers.name}}</v>
            </c>
            <c r=""B4"" t=""str"">
                <v>{{managers.department}}</v>
            </c>
            <c r=""C4"" t=""str"">
                <v>{{title}}</v>
            </c>			
        </row>
        <row r=""5"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A5"" t=""str"">
                <v>Employees</v>
            </c>
        </row>
        <row r=""6"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A6"" t=""str"">
                <v>{{employees.name}}</v>
            </c>
            <c r=""B6"" t=""str"">
                <v>{{employees.department}}</v>
            </c>
        </row>
        <row r=""100"" spans=""1:2"" x14ac:dyDescent=""0.25"">
            <c r=""A100"" t=""str"">
                <v>{{employees.name}}</v>
            </c>
            <c r=""B100"" t=""str"">
                <v>{{employees.department}}</v>
            </c>
        </row>		
    </sheetData>
    <mergeCells count=""1"">
        <mergeCell ref=""A2:B2"" />
    </mergeCells>
    <pageMargins left=""0.78749999999999998"" right=""0.78749999999999998"" top=""1.05277777777778"" bottom=""1.05277777777778"" header=""0.78749999999999998"" footer=""0.78749999999999998"" />
    <pageSetup orientation=""portrait"" useFirstPageNumber=""1"" horizontalDpi=""300"" verticalDpi=""300"" />
    <headerFooter>
        <oddHeader>&amp;C&amp;""Times New Roman,Regular""&amp;12&amp;A</oddHeader>
        <oddFooter>&amp;C&amp;""Times New Roman,Regular""&amp;12Page &amp;P</oddFooter>
    </headerFooter>
</worksheet>";