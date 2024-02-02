using System;
using System.Collections.Generic;
using System.IO;

namespace ArrayToExcel;

[Obsolete("Class is deprecated, please use ExcelBuilder instead.")]
public class ArrayToExcel
{
    [Obsolete("Method is deprecated, please use ExcelBuilder.Build instead.")]
    public static MemoryStream CreateExcel<T>(IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(items, schema);

    [Obsolete("Method is deprecated, please use ExcelBuilder.Build instead.")]
    public static void CreateExcel<T>(Stream stream, IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(stream, items, schema);
}

[Obsolete("Class is deprecated, please use CellText instead.")]
public class Text(string? value, bool wrap = false) : CellText(value, wrap);

[Obsolete("Class is deprecated, please use CellPercent instead.")]
public class Percent(object? value) : CellPercent(value);

[Obsolete("Class is deprecated, please use CellHyperlink instead.")]
public class Hyperlink(string link, string? text = null) : CellHyperlink(link, text)
{
    public Hyperlink(Uri link, string? text = null)
        : this(link.ToString(), text ?? link.OriginalString) { }

}

[Obsolete("Class is deprecated, please use CellFormula instead.")]
public class Formula(Func<uint, string> rowText) : CellFormula(rowText)
{
    public Formula(string text) : this(row => text) { }
}

