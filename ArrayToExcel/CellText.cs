using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ArrayToExcel;

public class CellText(string? value, bool wrapText = false) : ICellValue
{
    public virtual void Apply(Cell cell, uint row) => Apply(cell, value, wrapText);

    internal static void Apply(Cell cell, string? value, bool wrapText)
    {
        cell.InlineString = GetInlineString(value ?? string.Empty);
        cell.DataType = CellValues.InlineString;
        cell.StyleIndex = wrapText ? Styles.WrapText : Styles.Default;
    }

    static InlineString GetInlineString(string value)
    {
        return new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text(CellDefault.NormCellText(value))
        {
            Space = SpaceProcessingModeValues.Preserve
        });
    }
}
