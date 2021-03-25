using System;
using System.IO;

namespace MiniExcelLibs.Csv
{
    public class CsvConfiguration : IConfiguration
    {
	   public char Seperator { get; set; }
	   public Func<Stream, StreamReader> GetStreamReaderFunc { get; set; }
	   private static readonly CsvConfiguration _defaultConfiguration = new CsvConfiguration()
	   {
		  Seperator = ',',
		  GetStreamReaderFunc = (stream) => new StreamReader(stream)
	   };
	   internal static CsvConfiguration GetDefaultConfiguration() => _defaultConfiguration;
    }
}
