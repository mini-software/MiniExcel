using System;
using System.IO;
using System.Text;

namespace MiniExcelLibs.Csv
{
    public class CsvConfiguration : IConfiguration
    {
        private static Encoding _defaultEncoding = new UTF8Encoding(true);

        public char Seperator { get; set; } = ',';
        public string NewLine { get; set; } = "\r\n";
        public Func<Stream, StreamReader> StreamReaderFunc { get; set; } = (stream) => new StreamReader(stream, _defaultEncoding);
        public Func<Stream, StreamWriter> StreamWriterFunc { get; set; } = (stream) => new StreamWriter(stream, _defaultEncoding);

        internal readonly static CsvConfiguration DefaultConfiguration = new CsvConfiguration();
    }
}

