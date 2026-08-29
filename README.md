[![](https://img.shields.io/nuget/v/soenneker.excel.util.svg?style=for-the-badge)](https://www.nuget.org/packages/soenneker.excel.util/)
[![](https://img.shields.io/github/actions/workflow/status/soenneker/soenneker.excel.util/publish-package.yml?style=for-the-badge)](https://github.com/soenneker/soenneker.excel.util/actions/workflows/publish-package.yml)
[![](https://img.shields.io/nuget/dt/soenneker.excel.util.svg?style=for-the-badge)](https://www.nuget.org/packages/soenneker.excel.util/)
[![](https://img.shields.io/github/actions/workflow/status/soenneker/soenneker.excel.util/codeql.yml?label=CodeQL&style=for-the-badge)](https://github.com/soenneker/soenneker.excel.util/actions/workflows/codeql.yml)

# Soenneker.Excel.Util

Provides methods for reading and writing Excel files using strongly-typed objects with automatic property mapping and basic type conversion.

## Install

```bash
dotnet add package Soenneker.Excel.Util
```

## Quick start

```csharp
using Soenneker.Excel.Util.Registrars;
using Microsoft.Extensions.DependencyInjection;

var services = new ServiceCollection();
var result = services.AddExcelUtilAsSingleton();
```

Adds `IExcelUtil` as a singleton service.

## What you get

- `IExcelUtil` — Provides methods for reading and writing Excel files using strongly-typed objects with automatic property mapping and basic type conversion.
- `ExcelUtilRegistrar` — Provides methods for reading and writing Excel files using strongly-typed objects with automatic property mapping and basic type conversion.
- `ExcelColumnAttribute` — Represents the excel column attribute.

## API at a glance

| API | What it does | Result / important behavior |
| --- | --- | --- |
| `IExcelUtil.Read(filePath, sheetName)` | Reads data from an Excel worksheet and maps it to a list of objects of type `T`. | A list of objects of type `T` populated from the Excel worksheet. |
| `IExcelUtil.Write(objects, filePath, sheetName)` | Writes a list of objects to an Excel worksheet. | Returns no value; the requested change is complete when the method returns. |
| `ExcelUtilRegistrar.AddExcelUtilAsSingleton(services)` | Adds `IExcelUtil` as a singleton service. | The same service collection, so additional registrations can be chained. |
| `ExcelUtilRegistrar.AddExcelUtilAsScoped(services)` | Adds `IExcelUtil` as a scoped service. | The same service collection, so additional registrations can be chained. |
| `ExcelColumnAttribute.Name` | Gets name. | Gets name. |
