using System;
using System.IO;

namespace MiniExcelLibs.Csv
{
    public class CsvConfiguration : IConfiguration
    {
        public char Seperator { get; set; } = ',';
        public Func<Stream, StreamReader> GetStreamReaderFunc { get; set; } = (stream) => new StreamReader(stream);
    }
}

