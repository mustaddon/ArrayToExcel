using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ArrayToExcel;

public class CellDate : ICellValue
{
    public CellDate(DateTime? dateTime, bool dateOnly = false)
    {
        _value = dateTime?.ToString(_dateTimeFormat);
        _dateOnly = dateOnly;
    }

    public CellDate(DateTimeOffset? dateTime, bool dateOnly = false)
    {
        _value = dateTime?.ToString(_dateTimeFormat);
        _dateOnly = dateOnly;
    }

#if NET6_0_OR_GREATER
    public CellDate(DateOnly? date)
    {
        _value = date?.ToString(_dateFormat);
        _dateOnly = true;
    }

    static readonly string _dateFormat = "o";
#endif

    static readonly string _dateTimeFormat = "s";
    readonly string? _value;
    readonly bool _dateOnly;

    public virtual void Apply(Cell cell, uint row)
    {

        cell.CellValue = new(_value ?? string.Empty);
        cell.DataType = CellValues.Date;
        cell.StyleIndex = _dateOnly ? Styles.Date : Styles.DateTime;
    }
}
