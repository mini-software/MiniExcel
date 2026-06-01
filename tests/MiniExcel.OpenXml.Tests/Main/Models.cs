using MiniExcelLib.Tests.Common;

namespace MiniExcelLib.OpenXml.Tests.Main;

internal class AutoCheckType
{
    public Guid? @guid { get; set; }
    public bool? @bool { get; set; }
    public DateTime? datetime { get; set; }
    public string? @string { get; set; }
}

internal class CustomAttributesWihoutVaildPropertiesTestPoco
{
    [MiniExcelIgnore]
    public string? Test3 { get; set; }
    public string? Test5 { get; }
    public string? Test6 { get; private set; }
}

internal class Demo
{
    public string? Column1 { get; set; }
    public decimal Column2 { get; set; }
}

internal class DemoPocoHelloWorld
{
    public string? HelloWorld1 { get; set; }
}

internal class SaveAsControlChracterVO
{
    public string? Test { get; set; }
}

/// <summary>
/// Test class with multiple date and time properties using MiniExcelFormatAttribute
/// to verify that date/time formatting is correctly applied during Excel export.
/// </summary>
internal class DateTimeFormattingTestDto(
    DateTime shortDate,
    DateTime longDate,
    DateTime dateWithTime,
    TimeSpan timeOnly,
    DateTime isoDateTime,
    DateTime customDateTime,
    DateTime monthYear)
{
    /// <summary>
    /// Short date format (mm/dd/yyyy)
    /// </summary>
    [MiniExcelFormat("mm/dd/yyyy")]
    public DateTime ShortDate { get; set; } = shortDate;

    /// <summary>
    /// Long date format (dddd, mmmm dd, yyyy)
    /// </summary>
    [MiniExcelFormat("dddd, mmmm dd, yyyy")]
    public DateTime LongDate { get; set; } = longDate;

    /// <summary>
    /// Date with time format (yyyy-mm-dd hh:mm:ss)
    /// </summary>
    [MiniExcelFormat("yyyy-mm-dd hh:mm:ss")]
    public DateTime DateWithTime { get; set; } = dateWithTime;

    /// <summary>
    /// Time only format ([h]:mm:ss)
    /// </summary>
    [MiniExcelFormat("[h]:mm:ss")]
    public TimeSpan TimeOnly { get; set; } = timeOnly;

    /// <summary>
    /// ISO 8601 datetime format (yyyy-mm-ddThh:mm:ss)
    /// </summary>
    [MiniExcelFormat("yyyy-mm-dd\"T\"hh:mm:ss")]
    public DateTime IsoDateTime { get; set; } = isoDateTime;

    /// <summary>
    /// Custom European format (dd.mm.yyyy hh:mm)
    /// </summary>
    [MiniExcelFormat("dd.mm.yyyy hh:mm")]
    public DateTime CustomDateTime { get; set; } = customDateTime;

    /// <summary>
    /// Month and year format (mmmm yyyy)
    /// </summary>
    [MiniExcelFormat("mmmm yyyy")]
    public DateTime MonthYear { get; set; } = monthYear;
}


    /// <summary>
    /// Test class with multiple numeric properties using MiniExcelFormatAttribute
    /// to verify that formatting is correctly applied during Excel export.
    /// </summary>
internal class NumericFormattingTestDto(
    decimal currency,
    decimal alignedCurrency,
    decimal percentage,
    double scientificNotation,
    decimal fixedDecimal,
    long phoneNumber,
    long veryLongNumber,
    double customFormat)
{

    /// <summary>
    /// Regular currency format with 2 decimal places
    /// </summary>
    [MiniExcelFormat("\"$\"#,##0.00")]
    public decimal Currency { get; set; } = currency;

    /// <summary>
    /// Currency format with 2 decimal places, parentheses for negatives
    /// </summary>
    [MiniExcelFormat("$#,##0.00_);($#,##0.00)")]
    public decimal AlignedCurrency { get; set; } = alignedCurrency;

    /// <summary>
    /// Percentage format with 0 decimal places
    /// </summary>
    [MiniExcelFormat("0%")]
    public decimal Percentage { get; set; } = percentage;

    /// <summary>
    /// Scientific notation format with 2 decimal places
    /// </summary>
    [MiniExcelFormat("0.00E+00")]
    public double ScientificNotation { get; set; } = scientificNotation;

    [MiniExcelFormat("0.00E+00"), MiniExcelHidden]
    public double ScientificNotationDuplicate { get; set; } = scientificNotation;

    /// <summary>
    /// Fixed decimal places (6 decimal places)
    /// </summary>
    [MiniExcelFormat("0.000000")]
    public decimal FixedDecimal { get; set; } = fixedDecimal;

    /// <summary>
    /// Phone number format
    /// </summary>
    [MiniExcelFormat("[<=9999999]###-####;(###) ###-####")]
    public long PhoneNumber { get; set; } = phoneNumber;

    /// <summary>
    /// Simple integer format that shows the number in its full length (no scientific notation)
    /// </summary>
    [MiniExcelFormat("#")]
    public long VeryLongNumber { get; set; } = veryLongNumber;

    /// <summary>
    /// Simple decimal format with 3 decimal places
    /// </summary>
    [MiniExcelFormat("0.000")]
    public double CustomFormat { get; set; } = customFormat;
}

internal class UserAccount
{
    public Guid ID { get; set; }
    public string? Name { get; set; }
    public DateTime BoD { get; set; }
    public int Age { get; set; }
    public bool VIP { get; set; }
    public decimal Points { get; set; }
    public int IgnoredProperty => 1;
}

internal class ExcelAttributeDemo
{
    [MiniExcelColumnName("Column1")]
    public string? Test1 { get; set; }
    [MiniExcelColumnName("Column2")]
    public string? Test2 { get; set; }
    [MiniExcelIgnore]
    public string? Test3 { get; set; }
    [MiniExcelColumnIndex("I")] // system will convert "I" to 8 index
    public string? Test4 { get; set; }
    public string? Test5 { get; } //wihout set will ignore
    public string? Test6 { get; private set; } //un-public set will ignore
    [MiniExcelColumnIndex(3)] // start with 0
    public string? Test7 { get; set; }
}

internal class ExcelAttributeDemo2
{
    [MiniExcelColumn(Name = "Column1")]
    public string? Test1 { get; set; }
    [MiniExcelColumn(Name = "Column2")]
    public string? Test2 { get; set; }
    [MiniExcelColumn(Ignore = true)]
    public string? Test3 { get; set; }
    [MiniExcelColumn(IndexName = "I")] // system will convert "I" to 8 index
    public string? Test4 { get; set; }
    public string? Test5 { get; } //wihout set will ignore
    public string? Test6 { get; private set; } //un-public set will ignore
    [MiniExcelColumn(Index = 3)] // start with 0
    public string? Test7 { get; set; }
}
    
class LocalizationSupportDto(string firstName, string lastName, string address, int age)
{
    [MiniExcelColumn(Name = nameof(FirstName), ResourceType = typeof(Localization), Width = 15)]
    public string? FirstName { get; set; } = firstName;

    [MiniExcelColumn(Name = nameof(LastName), ResourceType = typeof(Localization), Width = 15)]
    public string? LastName { get; set; } = lastName;

    [MiniExcelColumnName("Address", ResourceType = typeof(Localization))]
    public string? Residency { get; set; } = address;

    [MiniExcelColumn(Name = nameof(Age), ResourceType = typeof(Localization), Width = 20)]
    public int Age { get; set; } = age;
}

internal class SaveAsFileWithDimensionByICollectionTestType
{
    public string? A { get; set; }
    public string? B { get; set; }
}
