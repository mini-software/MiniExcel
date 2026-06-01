namespace MiniExcelLib.Csv.Tests.Issues;

internal class Issue89Dto
{
    public WorkState State { get; set; }

    public enum WorkState
    {
        OnDuty,
        Leave,
        Fired
    }
}

internal class Issue142Dto
{
    [MiniExcelColumnName("CustomColumnName")]
    public string? MyProperty1 { get; set; }  //index = 1
    [MiniExcelIgnore]
    public string? MyProperty7 { get; set; } //index = null
    public string? MyProperty2 { get; set; } //index = 3
    [MiniExcelColumnIndex(6)]
    public string? MyProperty3 { get; set; } //index = 6
    [MiniExcelColumnIndex("A")] // equal column index 0
    public string? MyProperty4 { get; set; }
    [MiniExcelColumnIndex(2)]
    public string? MyProperty5 { get; set; } //index = 2
    public string? MyProperty6 { get; set; } //index = 4
}

internal class Issue142DuplicateColumnNameDto
{
    [MiniExcelColumnIndex("A")]
    public int MyProperty1 { get; set; }
    [MiniExcelColumnIndex("A")]
    public int MyProperty2 { get; set; }

    public int MyProperty3 { get; set; }
    [MiniExcelColumnIndex("B")]
    public int MyProperty4 { get; set; }
}

internal class Issue142OverIndexDto
{
    [MiniExcelColumnIndex("Z")]
    public int MyProperty1 { get; set; }
}

internal class Issue142ExcelColumnNameNotFoundDto
{
    [MiniExcelColumnIndex("B")]
    public int MyProperty1 { get; set; }
}

internal class Issue241Dto
{
    public string? Name { get; set; }

    [MiniExcelFormat("MM dd, yyyy")]
    public DateTime InDate { get; set; }
}

internal class Issue243Dto
{
    public string? Name { get; set; }
    public int Age { get; set; }
    public DateTime InDate { get; set; }
}

internal class TestIssue305Dto
{
    [MiniExcelFormat("yyyy-MM-dd")]
    public DateTimeOffset? Dt { get; set; }
}

internal class TestIssue312Dto
{
    [MiniExcelFormat("0,0.00")]
    public double? Value { get; set; }
}

internal class TestIssue316Dto
{
    public decimal Amount { get; set; }
    public DateTime CreateTime { get; set; }
}

internal class Issue507V01
{
    public string? A { get; set; }
    public DateTime B { get; set; }
    public string? C { get; set; }
    public int D { get; set; }
}

internal class Issue507V02
{
    public DateTime B { get; set; }
    public int D { get; set; }
}
