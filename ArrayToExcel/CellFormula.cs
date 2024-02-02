using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellFormula(Func<uint, string, string> cellText) : ICellValue
{
    public CellFormula(Func<uint, string> rowText) : this((row, cell) => rowText(row)) { }
    public CellFormula(string text) : this((row, col) => text) { }

    public virtual void Apply(Cell cell, uint row)
    {
        cell.CellFormula = new DocumentFormat.OpenXml.Spreadsheet.CellFormula(cellText(row, cell.CellReference!));
        cell.StyleIndex = 4;
    }
}
