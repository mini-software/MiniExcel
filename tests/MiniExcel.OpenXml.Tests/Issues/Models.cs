using System.ComponentModel;

namespace MiniExcelLib.OpenXml.Tests.Issues;

internal enum DescriptionEnum
{
    [Description("General User")] V1,
    [Description("General Administrator")] V2,
    [Description("Super Administrator")] V3
}

internal class DescriptionEnumDto
{
    public string? Name { get; set; }
    public DescriptionEnum? UserType { get; set; }
}

public class UserAccount
{
    public Guid ID { get; set; }
    public string? Name { get; set; }
    public DateTime BoD { get; set; }
    public int Age { get; set; }
    public bool VIP { get; set; }
    public decimal Points { get; set; }

    public int IgnoredProperty => 1;
}

internal class TestIssues133Dto
{
    public string? Id { get; set; }
    public string? Name { get; set; }
}

internal class Issue137Dto
{
    public double? 比例 { get; set; }
    public string? 商品 { get; set; }
    public int? 滿倉口數 { get; set; }
}

internal class Issue138Dto
{
    public DateTime? Date { get; set; }
    public int? 實單每日損益 { get; set; }
    public int? 程式每日損益 { get; set; }
    public string? 商品 { get; set; }
    public double? 滿倉口數 { get; set; }
    public double? 波段 { get; set; }
    public double? 當沖 { get; set; }
}

internal class Issue142Dto
{
    [MiniExcelColumnName("CustomColumnName")]
    public string? MyProperty1 { get; set; } //index = 1

    [MiniExcelIgnore] public string? MyProperty7 { get; set; } //index = null
    public string? MyProperty2 { get; set; } //index = 3
    [MiniExcelColumnIndex(6)] public string? MyProperty3 { get; set; } //index = 6

    [MiniExcelColumnIndex("A")] // equal column index 0
    public string? MyProperty4 { get; set; }

    [MiniExcelColumnIndex(2)] public string? MyProperty5 { get; set; } //index = 2
    public string? MyProperty6 { get; set; } //index = 4
}

internal class Issue142DtoVariant1
{
    [MiniExcelColumnIndex("Z")]
    public int MyProperty1 { get; set; }
}

internal class Issue142DtoVariant2
{
    [MiniExcelColumnIndex("B")]
    public int MyProperty1 { get; set; }
}

internal class TestIssue190Dto
{
    public int ID { get; set; }
    public string? Name { get; set; }
    public int Age { get; set; }
}

internal class TestIssue209Dto
{
    public int ID { get; set; }
    public string? Name { get; set; }
    public int SEQ { get; set; }
}

internal class Issue241Dto
{
    public string? Name { get; set; }

    [MiniExcelFormat("MM dd, yyyy")] public DateTime InDate { get; set; }
}

internal class Issue255DTO
{
    [MiniExcelFormat("yyyy")] public DateTime Time { get; set; }

    [MiniExcelColumn(Format = "yyyy")] public DateTime Time2 { get; set; }
}

internal class TestIssue280Dto
{
    [MiniExcelColumnWidth(20)] public int ID { get; set; }
    [MiniExcelColumnWidth(15.50)] public string? Name { get; set; }
}

internal class TestIssue286Dto
{
    public TestIssue286Enum E { get; set; }
}

internal enum TestIssue286Enum
{
    VIP1,
    VIP2
}

internal class TestIssue310Dto
{
    public int? V1 { get; set; }
}

internal class TestIssue312Dto
{
    [MiniExcelFormat("0,0.00")] public double? Value { get; set; }
}

internal class TestIssue331Dto
{
    public int Number { get; set; }
    public decimal DecimalNumber { get; set; }
    public double DoubleNumber { get; set; }
    public string? Text { get; set; }
}

internal class Issue409Dto
{
    public string? Units { get; set; }
    public double Quantity { get; set; }
}

internal class Issue422Enumerable(IEnumerable inner) : IEnumerable
{
    public int GetEnumeratorCount { get; private set; }

    public IEnumerator GetEnumerator()
    {
        GetEnumeratorCount++;
        return inner.GetEnumerator();
    }
}

internal class TestIssue430Dto
{
    [MiniExcelFormat("yyyy-MM-dd HH:mm:ss")]
    public DateTimeOffset Date { get; set; }
}

internal class Issue520Dto(long l1, DateTime dt, long l2)
{
    [MiniExcelColumn(Format = "R$ #,##0.00", Width = 15)]
    public long PaymentValue { get; set; } = l1;

    [MiniExcelColumn(Format = "dd/MM/yyyy", Width = 15)]
    public DateTime PaymentDate { get; set; } = dt;

    [MiniExcelColumn(Format = "R$ #,##0.00", Width = 15)]
    public long ValueToSettle { get; set; } = l2;
}

internal class Issue542
{
    [MiniExcelColumnIndex(0)] public Guid ID { get; set; }
    [MiniExcelColumnIndex(1)] public string? Name { get; set; }
}

internal class Issue585Variant1
{
    public string? Col1 { get; set; }
    public string? Col2 { get; set; }
    public string? Col3 { get; set; }
}

internal class Issue585Variant2
{
    public string? Col1 { get; set; }

    [MiniExcelColumnName("Col2")] public string? Prop2 { get; set; }

    public string? Col3 { get; set; }
}

internal class Issue585Variant3
{
    public string? Col1 { get; set; }

    [MiniExcelColumnIndex("B")] public string? Prop2 { get; set; }

    public string? Col3 { get; set; }
}

internal class TestIssueI4ZYUUDto
{
    [MiniExcelColumn(Name = "ID", Index = 0)]
    public string? MyProperty { get; set; }

    [MiniExcelColumn(Name = "CreateDate", Index = 1, Format = "yyyy-MM", Width = 100)]
    public DateTime MyProperty2 { get; set; }
}

internal class Issue658Dto
{
    public string? FirstName { get; set; }
    public string? LastName { get; set; }
}

internal class Issue697Dto
{
    public int First { get; set; }
    public int Second { get; set; }
    public int Third { get; set; }
    public int Fourth { get; set; }
}

internal class Issue869
{
    public string? Name { get; set; }
    public DateOnly? Date { get; set; }
}

internal class Issue880
{
    public string? Test { get; set; }
    public string? this[int i] => "";
}

internal class Issue888Dto
{
    public string? Key { get; set; }
    public string? Value { get; set; }
}

internal class Issue951Dto
{
    public string? Name { get; set; }
    public DateTime CreateDate { get; set; }
    public bool VIP { get; set; }
    public double Points { get; set; }

    public object this[string? test] => new();
}

internal class TestIssueI4YCLQ_2Dto
{
    [MiniExcelColumnIndex("A")] public string? 站点编码 { get; set; }
    [MiniExcelColumnIndex("B")] public string? 站址名称 { get; set; }
    [MiniExcelColumnIndex("C")] public string? 值1 { get; set; }
    [MiniExcelColumnIndex("D")] public string? 值2 { get; set; }
    [MiniExcelColumnIndex("E")] public string? 值3 { get; set; }
    [MiniExcelColumnIndex("F")] public string? 资源ID { get; set; }
    [MiniExcelColumnIndex("G")] public string? 值4 { get; set; }
    [MiniExcelColumnIndex("H")] public string? 值5 { get; set; }
    [MiniExcelColumnIndex("I")] public string? 值6 { get; set; }
    public string? 值7 { get; set; }
    [MiniExcelColumnName("NotExist")] public string? 值8 { get; set; }
}

internal class TestIssueI4WM67Dto
{
    public int ID { get; set; }
    public string? Name { get; set; }
}

internal class TestIssueI4TXGTDto
{
    public int ID { get; set; }
    public string? Name { get; set; }
    [DisplayName("Specification")] public string? Spc { get; set; }
    [DisplayName("Unit Price")] public decimal Up { get; set; }
}

internal class TestIssueI49RZHDto
{
    [MiniExcelFormat("dd-MM-yyyy")] public DateTime? dd { get; set; }
}

internal class TestIssueI40QA5Dto
{
    [MiniExcelColumnName(columnName: "EmployeeNo", aliases: new[] { "EmpNo", "No" })]
    public string? Empno { get; set; }

    public string? Name { get; set; }
}

internal class IssueI3X2ZLDTO
{
    public int Col1 { get; set; }
    public DateTime Col2 { get; set; }
}

internal class Issue149VO
{
    public string? Test { get; set; }
}
