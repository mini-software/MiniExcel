namespace MiniExcelLibs.Benchmarks;

public abstract class BenchmarkBase
{
    public const string FilePath = "Test1,000,000x10.xlsx";
    public const int RowCount = 1_000_000;

    public IEnumerable<DemoDto> GetValue() => Enumerable.Range(1, RowCount).Select(s => new DemoDto());

    public class DemoDto
    {
        public string Column1 { get; set; } = "Hello World";
        public string Column2 { get; set; } = "Hello World";
        public string Column3 { get; set; } = "Hello World";
        public string Column4 { get; set; } = "Hello World";
        public string Column5 { get; set; } = "Hello World";
        public string Column6 { get; set; } = "Hello World";
        public string Column7 { get; set; } = "Hello World";
        public string Column8 { get; set; } = "Hello World";
        public string Column9 { get; set; } = "Hello World";
        public string Column10 { get; set; } = "Hello World";
    }
}
