using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System;
using ArrayToExcel._internal;
using System.Collections.Generic;

namespace ArrayToExcel;

public class CellDefault(object? value) : ICellValue
{
    public void Apply(Cell cell, uint row) => Apply(cell, value);

    internal static void Apply(Cell cell, object? value)
    {
        cell.CellValue = GetCellValue(value);
        cell.DataType = GetCellType(value);
        cell.StyleIndex = cell.DataType == CellValues.Date ? 2 : 4u;
    }

    internal static CellValue GetCellValue(object? value)
    {
        if (value == null) return new();

        var type = value.GetType();

        if (type == typeof(bool))
            return new((bool)value ? _boolVals[1] : _boolVals[0]);

        if (type == typeof(DateTime))
            return new(((DateTime)value).ToString(_dateFormat, _cultureInfo));

        if (type == typeof(DateTimeOffset))
            return new(((DateTimeOffset)value).ToString(_dateFormat, _cultureInfo));

        if (type == typeof(double))
            return new(((double)value).ToString(_cultureInfo));

        if (type == typeof(decimal))
            return new(((decimal)value).ToString(_cultureInfo));

        if (type == typeof(float))
            return new(((float)value).ToString(_cultureInfo));

        return new(NormCellText(value.ToString()!));
    }


    internal static string NormCellText(string value)
    {
        return RegularExpressions.InvalidXmlChars().Replace(value.Length > _maxCellText ? value.Substring(0, _maxCellText) : value, string.Empty);
    }

    internal static CellValues GetCellType(object? value)
    {
        var type = value?.GetType() ?? typeof(object);

        if (type == typeof(bool))
            return CellValues.Boolean;

        if (_numericTypes.Contains(type))
            return CellValues.Number;

        if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
            return CellValues.Date;

        return CellValues.String;
    }

    static readonly HashSet<Type> _numericTypes = [
        typeof(short),
        typeof(ushort),
        typeof(int),
        typeof(uint),
        typeof(long),
        typeof(ulong),
        typeof(double),
        typeof(decimal),
        typeof(float)];

    static readonly CultureInfo _cultureInfo = CultureInfo.GetCultureInfo("en-US");
    static readonly string _dateFormat = "s";
    static readonly string[] _boolVals = ["0", "1"];

    const int _maxCellText = 32767;
}
