<Query Kind="Program">
  <NuGetReference>AngleSharp</NuGetReference>
  <NuGetReference>Dapper</NuGetReference>
  <NuGetReference>DocumentFormat.OpenXml</NuGetReference>
  <NuGetReference>Newtonsoft.Json</NuGetReference>
  <NuGetReference>System.Data.SqlClient</NuGetReference>
  <Namespace>Dapper</Namespace>
  <Namespace>MiniExcel</Namespace>
  <Namespace>Newtonsoft.Json</Namespace>
  <Namespace>System.Data.SqlClient</Namespace>
  <Namespace>System.IO.Compression</Namespace>
  <Namespace>System.Net.Http</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
  <Namespace>System.IO.Packaging</Namespace>
</Query>

void Main()
{
	//Read();
	Create();
}

void Read()
{
	//
	{
		// idea from : [Reading and Writing to Excel 2007 or Excel 2010 from C# - Part II: Basics | Robert MacLean](https://www.sadev.co.za/content/reading-and-writing-excel-2007-or-excel-2010-c-part-ii-basics)
		Package xlsxPackage = Package.Open(@"D:\git\MiniExcel\Samples\Xlsx\OpenXmlSDK_InsertCellValues\OpenXmlSDK_InsertCellValues.xlsx", FileMode.Open, FileAccess.ReadWrite);
		var allParts = xlsxPackage.GetParts().ToList();
		PackagePart worksheetPart = (from part in allParts
										 where part.ContentType.Equals("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
										 select part).FirstOrDefault();
		XElement worksheet = XElement.Load(XmlReader.Create(worksheetPart.GetStream()));

	}
}

void Create()
{
	var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.xlsx");
	Console.WriteLine(path);

	//MiniExcelHelper.CreateEmptyFie(path);
	//MiniExcelHelper.Create(path, new[] {1,2,3,4,5});
	MiniExcelHelper.Create(path, new[] {
		new { a = @"""<>+-*//}{\\n", b = 1234567890,c = true,d=DateTime.Now },
		new { a = "<test>Hello World</test>", b = -1234567890,c=false,d=DateTime.Now.Date}
	});

	// TODO: Dapper Row

	// TODO: Dictionary

	// TODO: Datatable
	
	
	//ProcessStartInfo psi = new ProcessStartInfo
	//{
	//	FileName = path,
	//	UseShellExecute = true
	//};
	//Process.Start(psi);
}

namespace MiniExcel
{
	using System;
	using System.IO;
	using System.IO.Compression;
	using System.Text;
	using MiniExcel;

	public static class MiniExcelHelper
	{
		public static Dictionary<string, object> GetDefaultFilesTree()
		{
			return new Dictionary<string, object>()
			{
				{"[Content_Types].xml",DefualtXml.defaultContent_TypesXml},
				{@"_rels\.rels",DefualtXml.defaultRels},
				{@"xl\_rels\workbook.xml.rels",DefualtXml.defaultWorkbookXmlRels},
				{@"xl\styles.xml",DefualtXml.defaultStylesXml},
				{@"xl\workbook.xml",DefualtXml.defaultWorkbookXml},
				{@"xl\worksheets\sheet1.xml",DefualtXml.defaultSheetXml},
			};
		}

		public static void Create(string path, object value, string startCell = "A1", bool printHeader = true)
		{
			var xy = Helper.ConvertCellToXY(startCell);

			var filesTree = GetDefaultFilesTree();
			{
				var sb = new StringBuilder();

				var yIndex = xy.Item2;

				if (value is System.Collections.ICollection)
				{
					var _vs = value as System.Collections.ICollection;
					object firstValue = null;
					{
						foreach (var v in _vs)
						{
							firstValue = v;
							break;
						}
					}
					var type = firstValue.GetType();
					var props = type.GetProperties();
					if (printHeader)
					{
						sb.AppendLine($"<x:row>");
						var xIndex = xy.Item1;
						foreach (var p in props)
						{
							var columname = Helper.ConvertXyToCell(xIndex, yIndex);
							sb.Append($"<x:c r=\"{columname}\" t=\"str\">");
							sb.Append($"<x:v>{p.Name}");
							sb.Append($"</x:v>");
							sb.Append($"</x:c>");
							xIndex++;
						}
						sb.AppendLine($"</x:row>");
						yIndex++;
					}

					foreach (var v in _vs)
					{
						sb.AppendLine($"<x:row>");
						var xIndex = xy.Item1;
						foreach (var p in props)
						{
							var cellValue = p.GetValue(v);
							var cellValueStr = Helper.GetValue(cellValue);
							var t = "t=\"str\"";
							{
								if (decimal.TryParse(cellValueStr, out var outV))
									t = "t=\"n\"";
								if (cellValue is bool)
								{
									t = "t=\"b\"";
									cellValueStr = (bool)cellValue ? "1" : "0";
								}
								if (cellValue is DateTime || cellValue is DateTime?)
								{
									t = "s=\"1\"";
									cellValueStr = ((DateTime)cellValue).ToOADate().ToString();
								}
							}
							var columname = Helper.ConvertXyToCell(xIndex, yIndex);
							sb.Append($"<x:c {t}>");
							sb.Append($"<x:v>{cellValueStr}");
							sb.Append($"</x:v>");
							sb.Append($"</x:c>");
							xIndex++;
						}
						sb.AppendLine($"</x:row>");
						yIndex++;
					}
				}
				filesTree[@"xl\worksheets\sheet1.xml"] = $@"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
<x:sheetData>{sb.ToString()}</x:sheetData>
</x:worksheet>";
			}
			CreateZipFileStream(path, filesTree);
		}

		public static void CreateEmptyFie(string path)
		{
			CreateZipFileStream(path, GetDefaultFilesTree());
		}

		private static void CreateStringEntry(ZipArchive archive, string entryPath, string content)
		{
			ZipArchiveEntry entry = archive.CreateEntry(entryPath);
			using (var zipStream = entry.Open())
			{
				var bytes = Encoding.ASCII.GetBytes(content);
				zipStream.Write(bytes, 0, bytes.Length);
			}
		}

		private static FileStream CreateZipFileStream(string path, Dictionary<string, object> filesTree)
		{
			using (FileStream stream = new FileStream(path, FileMode.CreateNew))
			{
				using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create))
				{
					foreach (var fileTree in filesTree)
					{
						ZipArchiveEntry entry = archive.CreateEntry(fileTree.Key);
						using (var zipStream = entry.Open())
						{
							var bytes = Encoding.ASCII.GetBytes(fileTree.Value.ToString());
							zipStream.Write(bytes, 0, bytes.Length);
						}
					}
				}
				return stream;
			}
		}

		internal static class DefualtXml
		{
			internal const string defaultRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""/xl/workbook.xml"" Id=""Rfc2254092b6248a9"" />
</Relationships>";

			internal const string defaultSheetXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:worksheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheetData>
    </x:sheetData>
</x:worksheet>";
			internal const string defaultWorkbookXmlRels = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"" Target=""/xl/worksheets/sheet1.xml"" Id=""R1274d0d920f34a32"" />
    <Relationship Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"" Target=""/xl/styles.xml"" Id=""R3db9602ace774fdb"" />
</Relationships>";

			internal const string defaultStylesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:styleSheet xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:fonts>
        <x:font />
    </x:fonts>
    <x:fills>
        <x:fill />
    </x:fills>
    <x:borders>
        <x:border />
    </x:borders>
    <x:cellStyleXfs>
        <x:xf />
    </x:cellStyleXfs>
    <x:cellXfs>
        <x:xf />
        <x:xf numFmtId=""14"" applyNumberFormat=""1"" />
    </x:cellXfs>
</x:styleSheet>";

			internal const string defaultWorkbookXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<x:workbook xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""
    xmlns:x=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
    <x:sheets>
        <x:sheet xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" name=""Sheet1"" sheetId=""1"" r:id=""R1274d0d920f34a32"" />
    </x:sheets>
</x:workbook>";

			internal const string defaultContent_TypesXml = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"" />
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml"" />
    <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"" />
    <Override PartName=""/xl/styles.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"" />
</Types>";
		}

		internal static class Helper
		{
			internal static string GetValue(object value) => value == null ? "" : value.ToString().Replace("<", "&lt;").Replace(">", "&gt;");

			/// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
			internal static string ConvertXyToCell(Tuple<int, int> xy)
			{
				return ConvertXyToCell(xy.Item1, xy.Item2);
			}

			/// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
			internal static string ConvertXyToCell(int x, int y)
			{
				int dividend = x;
				string columnName = String.Empty;
				int modulo;

				while (dividend > 0)
				{
					modulo = (dividend - 1) % 26;
					columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
					dividend = (int)((dividend - modulo) / 26);
				}
				return $"{columnName}{y}";
			}

			/// <summary>X=CellLetter,Y=CellNumber,ex:A1=(1,1),B2=(2,2)</summary>
			internal static Tuple<int, int> ConvertCellToXY(string cell)
			{
				const string keys = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";
				const int mode = 26;

				var x = 0;
				var cellLetter = GetCellLetter(cell);
				//AA=27,ZZ=702
				for (int i = 0; i < cellLetter.Length; i++)
					x = x * mode + keys.IndexOf(cellLetter[i]);

				var cellNumber = GetCellNumber(cell);
				return Tuple.Create(x, int.Parse(cellNumber));
			}

			internal static string GetCellNumber(string cell)
			{
				string cellNumber = string.Empty;
				for (int i = 0; i < cell.Length; i++)
				{
					if (Char.IsDigit(cell[i]))
						cellNumber += cell[i];
				}
				return cellNumber;
			}

			internal static string GetCellLetter(string cell)
			{
				string GetCellLetter = string.Empty;
				for (int i = 0; i < cell.Length; i++)
				{
					if (Char.IsLetter(cell[i]))
						GetCellLetter += cell[i];
				}
				return GetCellLetter;
			}
		}

	}
}
