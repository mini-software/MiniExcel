namespace MiniExcelLib.Csv.Tests.Main;

internal class TestDto
{
    public string? C1 { get; set; }
    public string? C2 { get; set; }
}

internal class CsvFieldMappingTest
{
    [MiniExcelColumnName("Column1")]
    public string? Test1;

    [MiniExcelColumnName("Column2")]
    public int Test2;

    [MiniExcelColumnIndex(0)]
    public decimal Test;
}

internal class MixedFieldPropertyTest
{
    [MiniExcelColumnName("F1")]
    public string? Field1;

    [MiniExcelColumnName("P1")]
    public string? Prop1 { get; set; }
}

internal class CsvFieldsWithoutAttributeDemo
{
    public string? NotMappedField;

    [MiniExcelColumnName("Mapped")]
    public string? MappedField;
}

internal class TestWithAlias
{
    [MiniExcelColumnName(columnName: "c1", aliases: ["column1", "col1"])]
    public string? C1 { get; set; }

    [MiniExcelColumnName(columnName: "c2", aliases: ["column2", "col2"])]
    public string? C2 { get; set; }
}