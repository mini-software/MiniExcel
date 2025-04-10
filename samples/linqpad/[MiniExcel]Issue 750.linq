<Query Kind="Program">
  <Namespace>System.IO.Compression</Namespace>
</Query>

void Main(){
	Issue750.Main();
}

public class Issue750
{
	public static void Main()
	{
		var templatePath = @"D:\git\MiniExcel\samples\xlsx\TestIssue20250403_SaveAsByTemplate_OPT.xlsx";
		var path = Path.GetTempPath() + Guid.NewGuid() + ".xlsx";
		Stream _outputFileStream;
		Console.WriteLine(path);
		using (var stream = File.Create(path))
		using (var templateStream = FileHelper.OpenSharedRead(templatePath))
		{
			_outputFileStream = stream;
			templateStream.CopyTo(_outputFileStream);
			var outputZipFile = new MiniExcelZipArchive(_outputFileStream, mode: ZipArchiveMode.Update, true, Encoding.UTF8);

			var templateSheets = outputZipFile.Entries
				.Where(w =>
					w.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
					w.FullName.StartsWith("/xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase))
				.ToList();
			var templateSheet = templateSheets[0];
			var fullName = templateSheet.FullName;
			var outputSheetStream = outputZipFile.CreateEntry(fullName);
			using (var outputZipSheetStream = templateSheet.Open())
			{
				var doc = new XmlDocument();
				doc.Load(outputZipSheetStream);
				outputZipSheetStream.Dispose();



				templateSheet.Delete(); // ZipArchiveEntry can't update directly, so need to delete then create logic




				XmlNamespaceManager _ns;
				_ns = new XmlNamespaceManager(new NameTable());
				_ns.AddNamespace("x", Config.SpreadsheetmlXmlns);
				_ns.AddNamespace("x14ac", Config.SpreadsheetmlXml_x14ac);

				var worksheet = doc.SelectSingleNode("/x:worksheet", _ns);
				var sheetData = doc.SelectSingleNode("/x:worksheet/x:sheetData", _ns);
				var newSheetData = sheetData?.Clone(); //avoid delete lost data

				var rows = newSheetData?.SelectNodes("x:row", _ns);


				sheetData.RemoveAll();
				sheetData.InnerText = "{{{{{{split}}}}}}"; //TODO: bad code smell

				var contents = doc.InnerXml.Split(new[] { $"<sheetData>{{{{{{{{{{{{split}}}}}}}}}}}}</sheetData>" }, StringSplitOptions.None);

				using (var writer = new StreamWriter(stream, Encoding.UTF8))
				{
					writer.Write(contents[0]);
					writer.Write($"<sheetData>"); // prefix problem

					for (int i = 0; i < 10000000; i++)
					{
						writer.Write($"{Guid.NewGuid}{Guid.NewGuid}{Guid.NewGuid}{Guid.NewGuid}{Guid.NewGuid}");
					}
					writer.Write($"<sheetData>");
				}
			}
		}

	}
	internal class Config
	{
		public const string SpreadsheetmlXmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		public const string SpreadsheetmlXmlStrictns = "http://purl.oclc.org/ooxml/spreadsheetml/main";
		public const string SpreadsheetmlXmlRelationshipns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		public const string SpreadsheetmlXmlStrictRelationshipns = "http://purl.oclc.org/ooxml/officeDocument/relationships";
		public const string SpreadsheetmlXml_x14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
	}
	// You can define other methods, fields, classes and namespaces here
	public class MiniExcelZipArchive : ZipArchive
	{
		public MiniExcelZipArchive(Stream stream, ZipArchiveMode mode, bool leaveOpen, Encoding entryNameEncoding)
		   : base(stream, mode, leaveOpen, entryNameEncoding)
		{
		}

		public new void Dispose()
		{
			Dispose(disposing: true);
			GC.SuppressFinalize(this);
		}
	}

	internal static partial class FileHelper
	{
		public static FileStream OpenSharedRead(string path) => File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
	}
}