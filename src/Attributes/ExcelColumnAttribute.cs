using System;

namespace Soenneker.Excel.Util.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public sealed class ExcelColumnAttribute(string name) : Attribute
{
    public string Name { get; } = name;
}