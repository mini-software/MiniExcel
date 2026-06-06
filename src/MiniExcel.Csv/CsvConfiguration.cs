using MiniExcelLib.Core;

namespace MiniExcelLib.Csv;

public class CsvConfiguration : MiniExcelBaseConfiguration
{
    private static readonly Encoding DefaultEncoding = new UTF8Encoding(true);

    internal static CsvConfiguration Default => new();
    
    public char Seperator { get; set; } = ',';
    public string NewLine { get; set; } = "\r\n";
    public bool ReadLineBreaksWithinQuotes { get; set; } = true;
    public bool ReadEmptyStringAsNull { get; set; } = false;
    public bool AlwaysQuote { get; set; } = false;
    public bool QuoteWhitespaces { get; set; } = true;
    public Func<string, string[]>? SplitFn { get; set; }

    // we leave the stream open by default and close it in the CsvReader unless the consumer decides to keep it open.
    public Func<Stream, StreamReader> StreamReaderFunc { get; set; } = stream => new StreamReader(stream, DefaultEncoding, true, 1024, true);
    public Func<Stream, StreamWriter> StreamWriterFunc { get; set; } = stream => new StreamWriter(stream, DefaultEncoding);
}
