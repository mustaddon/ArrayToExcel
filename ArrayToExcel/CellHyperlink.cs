using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellHyperlink(string link, string? text = null) : ICellValue
{
    public CellHyperlink(Uri link, string? text = null)
        : this(link.ToString(), text ?? link.OriginalString) { }

    public void Apply(Cell cell, uint row)
    {
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(_format.Value);
        cell.StyleIndex = 3;
    }

    readonly Lazy<string> _format = new(() => Format(link, text));

    static string Format(string link, string? text)
    {
        return string.Format(string.IsNullOrWhiteSpace(text) || text == link
            ? "HYPERLINK(\"{0}\")"
            : "HYPERLINK(\"{0}\",\"{1}\")",
            Fix(link), Fix(text));
    }

    static string? Fix(string? val) => val?.Replace("\"", "\"\"");
}
