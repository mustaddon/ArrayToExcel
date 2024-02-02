using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using System;
using ArrayToExcel._internal;
using System.Collections.Generic;

namespace ArrayToExcel;

public class CellDefault(object? value) : ICellValue
{
    public virtual void Apply(Cell cell, uint row) => Apply(cell, value);

    internal static void Apply(Cell cell, object? value)
    {
        cell.CellValue = GetCellValue(value);
        cell.DataType = GetCellType(value);
        cell.StyleIndex = Styles.Default;
    }

    static CellValue GetCellValue(object? value)
    {
        if (value == null) return new();

        if (value is bool boolVal)
            return new(_boolVals[boolVal ? 1 : 0]);

        if (value is double doubleVal)
            return new(doubleVal.ToString(_cultureInfo));

        if (value is decimal decimalVal)
            return new(decimalVal.ToString(_cultureInfo));

        if (value is float floatVal)
            return new(floatVal.ToString(_cultureInfo));

        return new(NormCellText(value.ToString()!));
    }


    internal static string NormCellText(string value)
    {
        return RegularExpressions.InvalidXmlChars().Replace(value.Length > _maxCellText ? value.Substring(0, _maxCellText) : value, string.Empty);
    }

    static CellValues GetCellType(object? value)
    {
        if (value == null)
            return CellValues.String;

        if (value is bool)
            return CellValues.Boolean;

        if (_numericTypes.Contains(value.GetType()))
            return CellValues.Number;

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
    static readonly string[] _boolVals = ["0", "1"];

    const int _maxCellText = 32767;
}
