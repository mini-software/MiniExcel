using System;
using System.IO;

namespace MiniExcelLibs.Csv
{
    public class CsvConfiguration : IConfiguration
    {
        public char Seperator { get; set; } = ',';
        public string NewLine { get; set; } = "\r\n";
        public Func<Stream, StreamReader> GetStreamReaderFunc { get; set; } = (stream) => new StreamReader(stream);

        internal readonly static CsvConfiguration DefaultConfiguration = new CsvConfiguration();
    }
}

