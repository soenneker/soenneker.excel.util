using Soenneker.Excel.Util.Abstract;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Reflection;
using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using Soenneker.Excel.Util.Attributes;
using Soenneker.Extensions.String;
using Soenneker.Extensions.Type;

namespace Soenneker.Excel.Util;

/// <inheritdoc cref="IExcelUtil"/>
public sealed class ExcelUtil : IExcelUtil
{
    private readonly ILogger<ExcelUtil> _logger;

    private static readonly ConcurrentDictionary<Type, PropertyInfo[]> _propertyCache = new();

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

        var headers = new List<string>();
        foreach (IXLCell cell in headerRow.Cells())
            headers.Add(cell.GetValue<string>());

        // Build header -> index map once (IndexOf inside the per-property loop is O(n*m)).
        var headerIndex = new Dictionary<string, int>(headers.Count, StringComparer.Ordinal);
        for (var i = 0; i < headers.Count; i++)
        {
            string h = headers[i];
            if (!headerIndex.ContainsKey(h))
                headerIndex.Add(h, i);
        }

        var skippedHeader = false;
        foreach (IXLRow? dataRow in worksheet.RowsUsed())
        {
            if (!skippedHeader)
            {
                skippedHeader = true;
                continue;
            }

            var obj = new T();

            foreach (PropertyInfo property in properties)
            {
                string headerName = property.GetCustomAttribute<ExcelColumnAttribute>()?.Name ?? property.Name;
                if (headerIndex.TryGetValue(headerName, out int colIndex))
                {
                    IXLCell? cell = dataRow.Cell(colIndex + 1);
                    var cellValue = cell.GetValue<string>();

                    if (!cellValue.IsNullOrWhiteSpace())
                    {
                        object? converted = property.PropertyType.ConvertPropertyValue(cellValue);
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
        return _propertyCache.GetOrAdd(type, static t => t.GetProperties(BindingFlags.Public | BindingFlags.Instance));
    }
}