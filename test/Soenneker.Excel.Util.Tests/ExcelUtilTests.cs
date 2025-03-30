using Soenneker.Excel.Util.Abstract;
using Soenneker.Tests.FixturedUnit;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using Soenneker.Excel.Util.Tests.Dtos;
using Xunit;

namespace Soenneker.Excel.Util.Tests;

[Collection("Collection")]
public class ExcelUtilTests : FixturedUnitTest
{
    private readonly IExcelUtil _excelUtil;

    public ExcelUtilTests(Fixture fixture, ITestOutputHelper output) : base(fixture, output)
    {
        _excelUtil = Resolve<IExcelUtil>(true);
    }

    [Fact]
    public void Default()
    {

    }

    [Fact]
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
