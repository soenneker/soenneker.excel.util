using System;

namespace Soenneker.Excel.Util.Attributes;

/// <summary>
/// Overrides the worksheet header used for a mapped property.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute(string name) : Attribute
{
    /// <summary>
    /// Gets the exact worksheet header name.
    /// </summary>
    public string Name { get; } = name;
}
