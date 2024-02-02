using DocumentFormat.OpenXml.Spreadsheet;

namespace ArrayToExcel;

public class CellPercent(object? value) : ICellValue
{
    public void Apply(Cell cell, uint row)
    {
        cell.CellValue = CellDefault.GetCellValue(value);
        cell.DataType = CellDefault.GetCellType(value);
        cell.StyleIndex = 5;
    }
}
