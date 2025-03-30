using Soenneker.Excel.Util.Attributes;

namespace Soenneker.Excel.Util.Tests.Dtos;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }

    [ExcelColumn("Email Address")]
    public string Email { get; set; } = string.Empty;
}