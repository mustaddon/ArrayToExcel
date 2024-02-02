using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ArrayToExcel;

public class CellText(string? value, bool wrap = false) : ICellValue
{
    public virtual void Apply(Cell cell, uint row) => Apply(cell, value, wrap);

    internal static void Apply(Cell cell, string? value, bool wrap)
    {
        cell.InlineString = GetInlineString(value ?? string.Empty);
        cell.DataType = CellValues.InlineString;
        cell.StyleIndex = wrap ? 4 : 6u;
    }

    static InlineString GetInlineString(string value)
    {
        return new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(CellDefault.NormCellText(value))
        {
            Space = SpaceProcessingModeValues.Preserve
        });
    }
}
