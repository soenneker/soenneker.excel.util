using Soenneker.Excel.Util.Abstract;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using Soenneker.Excel.Util.Attributes;
using Soenneker.Extensions.String;

namespace Soenneker.Excel.Util;

/// <inheritdoc cref="IExcelUtil"/>
public class ExcelUtil : IExcelUtil
{
    private readonly ILogger<ExcelUtil> _logger;

    public ExcelUtil(ILogger<ExcelUtil> logger)
    {
        _logger = logger;
    }

    public List<T> Read<T>(string filePath, string sheetName = "Sheet1") where T : new()
    {
        _logger.LogDebug("%% EXCELUTIL: -- Reading Excel from {path} ...", filePath);

        var result = new List<T>();
        PropertyInfo[] properties = GetCachedProperties(typeof(T));

        using var workbook = new XLWorkbook(filePath);
        IXLWorksheet? worksheet = workbook.Worksheet(sheetName);
        IXLRow? headerRow = worksheet.FirstRowUsed();
        List<string> headers = headerRow.Cells().Select(c => c.GetValue<string>()).ToList();

        foreach (IXLRow? dataRow in worksheet.RowsUsed().Skip(1))
        {
            var obj = new T();

            foreach (PropertyInfo property in properties)
            {
                string headerName = property.GetCustomAttribute<ExcelColumnAttribute>()?.Name ?? property.Name;
                int colIndex = headers.IndexOf(headerName);

                if (colIndex >= 0)
                {
                    IXLCell? cell = dataRow.Cell(colIndex + 1);
                    var cellValue = cell.GetValue<string>();

                    if (!cellValue.IsNullOrWhiteSpace())
                    {
                        object? converted = ConvertPropertyValue(property.PropertyType, cellValue);
                        if (converted != null)
                            property.SetValue(obj, converted);
                    }
                }
            }

            result.Add(obj);
        }

        _logger.LogDebug("%% EXCELUTIL: -- Finished reading Excel");

        return result;
    }

    public void Write<T>(List<T> objects, string filePath, string sheetName = "Sheet1")
    {
        _logger.LogDebug("%% EXCELUTIL: -- Writing Excel to {path} ...", filePath);

        PropertyInfo[] properties = GetCachedProperties(typeof(T));

        using var workbook = new XLWorkbook();
        IXLWorksheet worksheet = workbook.Worksheets.Add(sheetName);

        // Write headers
        for (var i = 0; i < properties.Length; i++)
        {
            string header = properties[i].GetCustomAttribute<ExcelColumnAttribute>()?.Name ?? properties[i].Name;
            worksheet.Cell(1, i + 1).Value = header;
        }

        // Write data
        for (var rowIndex = 0; rowIndex < objects.Count; rowIndex++)
        {
            T obj = objects[rowIndex];
            for (var colIndex = 0; colIndex < properties.Length; colIndex++)
            {
                object? value = properties[colIndex].GetValue(obj);
                worksheet.Cell(rowIndex + 2, colIndex + 1).Value = value?.ToString() ?? string.Empty;
            }
        }

        workbook.SaveAs(filePath);

        _logger.LogDebug("%% EXCELUTIL: -- Finished writing Excel");
    }

    private static PropertyInfo[] GetCachedProperties(Type type)
    {
        return type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
    }

    private static object? ConvertPropertyValue(Type targetType, string value)
    {
        if (targetType == typeof(string))
            return value;

        if (Nullable.GetUnderlyingType(targetType) is Type underlying)
        {
            if (value.IsNullOrWhiteSpace())
                return null;

            targetType = underlying;
        }

        return targetType switch
        {
            Type t when t == typeof(int) && int.TryParse(value, out int i) => i,
            Type t when t == typeof(long) && long.TryParse(value, out long l) => l,
            Type t when t == typeof(short) && short.TryParse(value, out short s) => s,
            Type t when t == typeof(ushort) && ushort.TryParse(value, out ushort us) => us,
            Type t when t == typeof(uint) && uint.TryParse(value, out uint ui) => ui,
            Type t when t == typeof(ulong) && ulong.TryParse(value, out ulong ul) => ul,
            Type t when t == typeof(byte) && byte.TryParse(value, out byte b) => b,
            Type t when t == typeof(sbyte) && sbyte.TryParse(value, out sbyte sb) => sb,
            Type t when t == typeof(float) && float.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out float f) => f,
            Type t when t == typeof(double) && double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) => d,
            Type t when t == typeof(decimal) && decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec) => dec,
            Type t when t == typeof(bool) && bool.TryParse(value, out bool bo) => bo,
            Type t when t == typeof(DateTime) && DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dt) => dt,
            Type t when t == typeof(TimeSpan) && TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out TimeSpan ts) => ts,
            Type t when t == typeof(Guid) && Guid.TryParse(value, out Guid g) => g,
            Type t when t == typeof(Uri) && Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri) => uri,
            Type t when t.IsEnum && Enum.TryParse(t, value, ignoreCase: true, out object? e) => e,
            _ => null
        };
    }
}