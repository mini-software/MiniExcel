using MiniExcelLib.OpenXml.Models;

namespace MiniExcelLib.OpenXml.Tests.MultipleSheets;

[MiniExcelSheet(Name = "Users")]
internal class UserDto
{
    public string? Name { get; set; }
    public int Age { get; set; }
}

[MiniExcelSheet(Name = "Departments", State = SheetState.Hidden)]
internal class DepartmentDto
{
    public string? ID { get; set; }
    public string? Name { get; set; }
}