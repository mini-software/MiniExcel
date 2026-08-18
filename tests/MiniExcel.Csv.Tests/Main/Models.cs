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

internal class JaggedRowsMappingTest
{
    public string? C1 { get; init; }
    public string? C2 { get; init; }
    public string? C3 { get; init; }
    public string? C4 { get; init; }
    public string? C5 { get; init; }
    public string? C6 { get; init; }
    public string? C7 { get; init; }
    public string? C8 { get; init; }
    public string? C9 { get; init; }
    public string? C10 { get; init; }
    public string? C11 { get; init; }
    public string? C12 { get; init; }
    public string? C13 { get; init; }
    public string? C14 { get; init; }
    public string? C15 { get; init; }
    public string? C16 { get; init; }
    public string? C17 { get; init; }
    public string? C18 { get; init; }
    public string? C19 { get; init; }
    public string? C20 { get; init; }
    public string? C21 { get; init; }
    public string? C22 { get; init; }
    public string? C23 { get; init; }
    public string? C24 { get; init; }
    public string? C25 { get; init; }
    public string? C26 { get; init; }
    public string? C27 { get; init; }
    public string? C28 { get; init; }
    public string? C29 { get; init; }
    public string? C30 { get; init; }
    public string? C31 { get; init; }
    public string? C32 { get; init; }
    public string? C33 { get; init; }
    public string? C34 { get; init; }
}
