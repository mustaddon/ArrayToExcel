using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellFormula(Func<uint, string, string> value, bool wrapText = false) : ICellValue
{
    public CellFormula(Func<uint, string> value, bool wrapText = false) : this((row, cell) => value(row), wrapText) { }
    public CellFormula(string text, bool wrapText = false) : this((row, col) => text, wrapText) { }

    public virtual void Apply(Cell cell, uint row)
    {
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(value(row, cell.CellReference!));
        cell.StyleIndex = wrapText ? Styles.WrapText : Styles.Default;
    }
}
