using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellDate : ICellValue
{
    public CellDate(DateTime? value, bool dateOnly = false)
    {
        _apply = value == null
            ? x => Apply(x, string.Empty, dateOnly)
            : x => Apply(x, value.Value, dateOnly);
    }

    public CellDate(DateTimeOffset? value, bool dateOnly = false)
    {
        _apply = value == null 
            ? x => Apply(x, string.Empty, dateOnly) 
            : x => Apply(x, value.Value, dateOnly);
    }

    readonly Action<Cell> _apply;

    public virtual void Apply(Cell cell, uint row) => _apply(cell);

    internal static void Apply(Cell cell, DateTimeOffset value, bool dateOnly)
        => Apply(cell, value.ToString(_dateTimeFormat), dateOnly);

    internal static void Apply(Cell cell, DateTime value, bool dateOnly)
        => Apply(cell, value.ToString(_dateTimeFormat), dateOnly);

    static readonly string _dateTimeFormat = "s";

    static void Apply(Cell cell, string value, bool dateOnly)
    {
        cell.CellValue = new(value);
        cell.DataType = CellValues.Date;
        cell.StyleIndex = dateOnly ? Styles.Date : Styles.DateTime;
    }


#if NET6_0_OR_GREATER
    public CellDate(DateOnly? value)
    {
        _apply = value == null
            ? x => Apply(x, string.Empty, true)
            : x => Apply(x, value.Value);
    }

    static readonly string _dateFormat = "o";

    internal static void Apply(Cell cell, DateOnly value)
        => Apply(cell, value.ToString(_dateFormat), true);
#endif
}
