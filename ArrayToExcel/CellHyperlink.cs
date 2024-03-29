﻿using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellHyperlink(string link, string? text = null) : ICellValue
{
    public CellHyperlink(Uri link, string? text = null)
        : this(link.ToString(), text ?? link.OriginalString) { }

    readonly Lazy<string> _format = new(() => Format(link, text));

    public virtual void Apply(Cell cell, uint row) => Apply(cell, _format.Value);

    internal static void Apply(Cell cell, Uri link) => Apply(cell, Format(link.ToString(), link.OriginalString));

    static void Apply(Cell cell, string value)
    {
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(value);
        cell.StyleIndex = Styles.Hyperlink;
    }

    static string Format(string link, string? text)
    {
        return string.Format(string.IsNullOrWhiteSpace(text) || text == link
            ? "HYPERLINK(\"{0}\")"
            : "HYPERLINK(\"{0}\",\"{1}\")",
            Fix(link), Fix(text));
    }

    static string? Fix(string? val) => val?.Replace("\"", "\"\"");
}
