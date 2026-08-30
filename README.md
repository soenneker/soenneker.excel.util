[![](https://img.shields.io/nuget/v/soenneker.excel.util.svg?style=for-the-badge)](https://www.nuget.org/packages/soenneker.excel.util/)
[![](https://img.shields.io/github/actions/workflow/status/soenneker/soenneker.excel.util/publish-package.yml?style=for-the-badge)](https://github.com/soenneker/soenneker.excel.util/actions/workflows/publish-package.yml)
[![](https://img.shields.io/nuget/dt/soenneker.excel.util.svg?style=for-the-badge)](https://www.nuget.org/packages/soenneker.excel.util/)
[![](https://img.shields.io/github/actions/workflow/status/soenneker/soenneker.excel.util/codeql.yml?label=CodeQL&style=for-the-badge)](https://github.com/soenneker/soenneker.excel.util/actions/workflows/codeql.yml)

# Soenneker.Excel.Util

Maps rows in an `.xlsx` worksheet to strongly typed objects and writes object lists to a new workbook using ClosedXML.

## Install

```bash
dotnet add package Soenneker.Excel.Util
```

## Define a row type

```csharp
using Soenneker.Excel.Util.Attributes;

public sealed class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }

    [ExcelColumn("Email Address")]
    public string Email { get; set; } = "";
}
```

Public instance properties must have public getters and setters and cannot be indexers. By default the header must exactly match the property name. `ExcelColumnAttribute` supplies a different exact header.

## Register and use

```csharp
using Microsoft.Extensions.DependencyInjection;
using Soenneker.Excel.Util.Abstract;
using Soenneker.Excel.Util.Registrars;

var services = new ServiceCollection();
services.AddLogging();
services.AddExcelUtilAsSingleton();

using ServiceProvider provider = services.BuildServiceProvider();
IExcelUtil excel = provider.GetRequiredService<IExcelUtil>();

var people = new List<Person>
{
    new() { Name = "Alice", Age = 30, Email = "alice@example.com" },
    new() { Name = "Bob", Age = 25, Email = "bob@example.com" }
};

excel.Write(people, @"C:\exports\people.xlsx", "People");

List<Person> imported = excel.Read<Person>(@"C:\exports\people.xlsx", "People");
```

`AddExcelUtilAsSingleton()` and `AddExcelUtilAsScoped()` register the same stateless utility with different DI lifetimes. Each operation creates and disposes its own workbook; no workbook or file handle is retained by the service.

## Mapping behavior

- The first used row is treated as the header row. Each subsequent used row creates one `T`.
- Header matching is ordinal and case-sensitive. Extra columns are ignored; a missing mapped column or blank cell leaves the property's constructor/default value unchanged.
- Duplicate worksheet headers resolve to the first matching column.
- Writing uses the mapped public properties as columns and calls `ToString()` for non-null values, so cells are written as text. Reading gets cell values as text and applies the package's basic property conversion.
- An empty worksheet returns an empty list. A missing file or worksheet, an unreadable workbook, or a failed type conversion is reported by the underlying API.

`Write` creates a new workbook containing the requested sheet and overwrites the destination path; it does not append to or preserve an existing workbook, formatting, formulas, or other sheets. Ensure the parent directory exists and use a temporary file plus an atomic replace when a partially written destination is unacceptable.

Both APIs are synchronous and ClosedXML loads workbook data in memory. Apply file-size and decompression limits before accepting untrusted uploads, and avoid using this utility for very large streaming imports or exports.
