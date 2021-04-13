
namespace MiniExcelLibs.OpenXml
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using System.Xml;

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
		  var doc = new XmlDocument();
		  doc.LoadXml(sheetXml);

		  XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable);
		  ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

		  var worksheet = doc.SelectSingleNode("/x:worksheet", ns);
		  var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", ns);

		  // ==== Dimension ====
		  // note : dimension need to put on the top ![image](https://user-images.githubusercontent.com/12729184/114507911-5dd88400-9c66-11eb-94c6-82ed7bdb5aab.png)
		  var dimension = doc.SelectSingleNode("/x:worksheet/x:dimension", ns) as XmlElement;
		  // update dimension
		  if (dimension == null)
			 throw new NotImplementedException("Excel Dimension Xml is null, please issue file for me. https://github.com/shps951023/MiniExcel/issues");

		  if (dimension != null)
		  {
			 var rows = sheetData.SelectNodes($"x:row", ns);
			 var maxRowIndexDiff = 0;
			 foreach (XmlElement row in rows)
			 {
				IEnumerable ienumerable = null;

				foreach (XmlElement c in row.SelectNodes($"x:c", ns))
				{
				    var v = c.SelectSingleNode("x:v", ns);
				    if (v?.InnerText == null)
					   continue;

				    var matchs = (_isExpressionRegex.Matches(v.InnerText).Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value));
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
			 var @refs = dimension.GetAttribute("ref").Split(':');
			 var letter = new String(refs[1].Where(Char.IsLetter).ToArray());
			 var digit = int.Parse(new String(refs[1].Where(Char.IsDigit).ToArray()));

			 dimension.SetAttribute("ref", $"{refs[0]}:{letter}{digit + maxRowIndexDiff}");
			 //writer.Write(CleanXml(newDimension.OuterXml));
		  }

		  
		  var newSheetData = sheetData.Clone();
		  sheetData.RemoveAll();
		  sheetData.InnerText = "{{{{{{split}}}}}}";

		  var contents = doc.InnerXml.Split(new string[] { "<sheetData>{{{{{{split}}}}}}</sheetData>" }, StringSplitOptions.None); ;
		  using (var writer = new StreamWriter(stream, Encoding.UTF8))
		  {
			 writer.Write(contents[0]);

			 //Q.Why so complex? A.because try to avoid render row OOM
			 //render sheetData
			 writer.Write("<sheetData>");
			 int originRowIndex;
			 int rowIndexDiff = 0;
			 var rows = newSheetData.SelectNodes($"x:row", ns);
			 foreach (XmlElement row in rows)
			 {
				var rowCotainIEnumerable = false;
				IEnumerable ienumerable = null;
				List<string> propKeys = null;
				Type ienumerableGenricType = null;
				string ienumerableKey = null;

				//TODO: some xlsx without r
				originRowIndex = int.Parse(row.GetAttribute("r"));

				//TODO: need to remove namespace

				// check if contains IEnumerble row
				{
				    var cs = row.SelectNodes($"x:c", ns);
				    foreach (XmlElement c in cs)
				    {
					   var cr = c.GetAttribute("r");
					   var letter = new String(cr.Where(Char.IsLetter).ToArray());
					   c.SetAttribute("r", $"{letter}{{{{{{MiniExcel_RowIndex}}}}}}");

					   var v = c.SelectSingleNode("x:v", ns);
					   if (v?.InnerText == null)
						  continue;

					   var matchs = (_isExpressionRegex.Matches(v.InnerText).Cast<Match>().GroupBy(x => x.Value).Select(varGroup => varGroup.First().Value));
					   foreach (var item in matchs)
					   {
						  var keys = item.Split('.');
						  var value = inputMaps[keys[0]];
						  if (value is IEnumerable && !(value is string))
						  {


							 if (propKeys == null)
								propKeys = new List<string>();
							 propKeys.Add(keys[1]); //TODO: check if not contain 1 index
											    //only check first one match IEnumerable, so only render one collection at same row
							 if (rowCotainIEnumerable == false)
							 {
								ienumerableKey = keys[0];
								// get ienumerable runtime type
								foreach (var element in value as IEnumerable)
								{
								    if (element != null)
								    {
									   ienumerableGenricType = element.GetType();
									   break;
								    }
								}

								ienumerable = value as IEnumerable;
								rowCotainIEnumerable = true;
							 }
						  }
						  else
						  {
							 v.InnerText = v.InnerText.Replace($"{{{{{keys[0]}}}}}", value.ToString()); //TODO: auto check type and set value
						  }
					   }
				    }
				}



				var newRowIndex = originRowIndex + rowIndexDiff;
				if (rowCotainIEnumerable && ienumerable != null)
				{
				    var first = true;
				    foreach (var element in ienumerable)
				    {
					   var newRow = row.Clone() as XmlElement;
					   newRow.SetAttribute("r", newRowIndex.ToString());
					   newRow.InnerXml = row.InnerXml.Replace($"{{{{{{MiniExcel_RowIndex}}}}}}", newRowIndex.ToString());

					   foreach (var key in propKeys)
					   {
						  var prop = ienumerableGenricType.GetProperty(key);
						  newRow.InnerXml = newRow.InnerXml.Replace($"{{{{{ienumerableKey}.{key}}}}}", prop.GetValue(element).ToString());
					   }

					   // note: only first time need add diff ![image](https://user-images.githubusercontent.com/12729184/114494728-6bceda80-9c4f-11eb-9685-8b5ed054eabe.png)
					   if (!first)
						  rowIndexDiff++;
					   first = false;

					   newRowIndex++;
					   writer.Write(CleanXml(newRow.OuterXml));
				    }
				}
				else
				{
				    row.SetAttribute("r", newRowIndex.ToString());
				    row.InnerXml = row.InnerXml.Replace($"{{{{{{MiniExcel_RowIndex}}}}}}", newRowIndex.ToString());
				    writer.Write(CleanXml(row.OuterXml));
				}

			 }
			 writer.Write("</sheetData>");
			 writer.Write(contents[1]);
		  }
	   }
	   private static string CleanXml(string xml)
	   {
		  //TODO: need to optimize
		  return xml.Replace("xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"", "").Replace("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"", "");
	   }
    }
}
