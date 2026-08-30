using System.Collections.Generic;

namespace Soenneker.Excel.Util.Abstract;

/// <summary>
/// Reads and writes Excel worksheets by mapping columns to public, readable and writable properties.
/// </summary>
public interface IExcelUtil
{
    /// <summary>
    /// Reads data from an Excel worksheet and maps it to a list of objects of type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The type of objects to map the Excel data to. Must have a parameterless constructor.</typeparam>
    /// <param name="filePath">The full path to the Excel file to read.</param>
    /// <param name="sheetName">The exact name of the worksheet to read. Defaults to <c>Sheet1</c>.</param>
    /// <returns>Objects populated from the worksheet rows after the header row, or an empty list when the worksheet is empty.</returns>
    List<T> Read<T>(string filePath, string sheetName = "Sheet1") where T : new();

    /// <summary>
    /// Writes a list of objects to an Excel worksheet.
    /// </summary>
    /// <typeparam name="T">The type of objects to write to the Excel file.</typeparam>
    /// <param name="objects">The list of objects to write.</param>
    /// <param name="filePath">The full path to the Excel file to create or overwrite.</param>
    /// <param name="sheetName">The name of the worksheet to create. Defaults to <c>Sheet1</c>.</param>
    void Write<T>(List<T> objects, string filePath, string sheetName = "Sheet1");
}
