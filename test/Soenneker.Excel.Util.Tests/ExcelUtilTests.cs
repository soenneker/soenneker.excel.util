using Soenneker.Excel.Util.Abstract;
using Soenneker.Tests.HostedUnit;
using System.Collections.Generic;
using System.IO;
using AwesomeAssertions;
using Soenneker.Excel.Util.Tests.Dtos;

namespace Soenneker.Excel.Util.Tests;

[ClassDataSource<Host>(Shared = SharedType.PerTestSession)]
public class ExcelUtilTests : HostedUnitTest
{
    private readonly IExcelUtil _excelUtil;

    public ExcelUtilTests(Host host) : base(host)
    {
        _excelUtil = Resolve<IExcelUtil>(true);
    }

    [Test]
    public void Default()
    {

    }

    [Test]
    public void Write_And_Read_ShouldPreserveData()
    {
        // Arrange
        var people = new List<Person>
        {
            new() { Name = "Alice", Age = 30, Email = "alice@example.com" },
            new() { Name = "Bob", Age = 25, Email = "bob@example.com" }
        };

        string filePath = Path.Combine(Path.GetTempPath(), $"test_{Path.GetRandomFileName()}.xlsx");

        try
        {
            // Act
            _excelUtil.Write(people, filePath);
            var readBack = _excelUtil.Read<Person>(filePath);

            // Assert
            readBack.Should().HaveCount(2);
            readBack[0].Name.Should().Be("Alice");
            readBack[0].Age.Should().Be(30);
            readBack[0].Email.Should().Be("alice@example.com");

            readBack[1].Name.Should().Be("Bob");
            readBack[1].Age.Should().Be(25);
            readBack[1].Email.Should().Be("bob@example.com");
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }
}
