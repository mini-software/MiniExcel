using System;
using System.IO;
using System.Text;

namespace MiniExcelLibs.Csv;

public class CsvConfiguration : MiniExcelConfiguration
{
    private static readonly Encoding DefaultEncoding = new UTF8Encoding(true);

    public char Seperator { get; set; } = ',';
    public string NewLine { get; set; } = "\r\n";
    public bool ReadLineBreaksWithinQuotes { get; set; } = true;
    public bool ReadEmptyStringAsNull { get; set; } = false;
    public bool AlwaysQuote { get; set; } = false;
    public bool QuoteWhitespaces { get; set; } = true;
    public Func<string, string[]>? SplitFn { get; set; }
    public Func<Stream, StreamReader> StreamReaderFunc { get; set; } = (stream) => new StreamReader(stream, DefaultEncoding);
    public Func<Stream, StreamWriter> StreamWriterFunc { get; set; } = (stream) => new StreamWriter(stream, DefaultEncoding);

    internal static readonly CsvConfiguration DefaultConfiguration = new CsvConfiguration();
}