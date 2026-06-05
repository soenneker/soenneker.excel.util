using System;

namespace Soenneker.Excel.Util.Attributes;

/// <summary>
/// Represents the excel column attribute.
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute(string name) : Attribute
{
    /// <summary>
    /// Gets name.
    /// </summary>
    public string Name { get; } = name;
}