namespace MiniExcelLib.Benchmarks;

public abstract class BenchmarkBase
{
    protected const string FilePath = "Test100,000x10.xlsx";
    protected const int RowCount = 100_000;

    protected IEnumerable<DemoDto> GetValue() => Enumerable.Range(1, RowCount).Select(_ => new DemoDto());

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
