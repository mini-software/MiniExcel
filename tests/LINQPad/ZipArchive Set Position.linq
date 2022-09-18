<Query Kind="Program">
  <NuGetReference>MiniExcel</NuGetReference>
  <Namespace>MiniExcelLibs</Namespace>
</Query>

void Main()
{
	ConsoleApplication.Program.Main(null);
}

// You can define other methods, fields, classes and namespaces here


namespace ConsoleApplication
{
	using System;
	using System.IO;
	using System.IO.Compression;
	public class Program
	{
		public static void Main(string[] args)
		{
			var path = Path.GetTempPath() + Guid.NewGuid() + ".zip";
			Console.WriteLine(path);
			using (FileStream zipToOpen = new FileStream(path, FileMode.Create))
			{
				using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
				{
					ZipArchiveEntry readmeEntry = archive.CreateEntry("Readme.txt");
					using (StreamWriter writer = new StreamWriter(readmeEntry.Open()))
					{
						writer.WriteLine("Information about this package.");
						writer.Flush();
						writer.BaseStream.Position=0;
						writer.WriteLine("========================");
					}
				}
			}
		}
	}
}