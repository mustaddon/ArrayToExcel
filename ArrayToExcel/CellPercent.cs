using DocumentFormat.OpenXml.Spreadsheet;

namespace ArrayToExcel;

public class CellPercent(object? value) : CellDefault(value)
{
    public override void Apply(Cell cell, uint row)
    {
        base.Apply(cell, row);
        cell.StyleIndex = Styles.Percentage;
    }
}
