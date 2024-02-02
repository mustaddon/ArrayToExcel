using DocumentFormat.OpenXml.Spreadsheet;

namespace ArrayToExcel;

public interface ICellValue
{
    void Apply(Cell cell, uint row);
}