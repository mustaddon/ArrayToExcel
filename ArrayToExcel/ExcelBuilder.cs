using ArrayToExcel._internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ArrayToExcel;


public class ExcelBuilder
{
    public static MemoryStream Build<T>(IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
    {
        var ms = new MemoryStream();
        Build(ms, items, schema);
        ms.Position = 0;
        return ms;
    }

    public static void Build<T>(Stream stream, IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
    {
        var builder = new SchemaBuilder<T>(items);
        schema?.Invoke(builder);
        CreateExcel(stream, new[] { builder.Schema }.Concat(builder.Childs));
    }

    public static bool DefaultWrapText = false;

    static void CreateExcel(Stream stream, IEnumerable<SheetSchema> sheetSchemas)
    {
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookpart = document.AddWorkbookPart();
        workbookpart.Workbook = new Workbook();

        var sheets = workbookpart.Workbook.AppendChild(new Sheets());

        AddStyles(workbookpart);

        var sheetNames = new HashSet<string>();
        var sheetId = 0u;

        foreach (var sheetSchema in sheetSchemas)
        {
            sheetId++;

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();

            var sheet = new Sheet()
            {
                Id = workbookpart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = NormSheetName(sheetSchema.SheetName, sheetId, sheetNames),
            };

            sheetNames.Add(sheet.Name.Value ?? string.Empty);
            sheets.AppendChild(sheet);

            var cols = worksheetPart.Worksheet.AppendChild(new Columns());
            var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            if (sheetSchema.Columns.Count == 0)
            {
                cols.AppendChild(new Column() { Min = 1, Max = 1, BestFit = true });
                continue;
            }

            cols.Append(sheetSchema.Columns.Select((x, i) => new Column() { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = x.Width, CustomWidth = true, BestFit = true }));

            sheetData.Append(GetRows(sheetSchema.Items, sheetSchema.Columns, sheetSchema.WrapText ?? DefaultWrapText));

            worksheetPart.Worksheet.AppendChild(new AutoFilter() { Reference = $"A1:{GetColReference(sheetSchema.Columns.Count - 1)}{sheetData.ChildElements.Count}" });
        }

        workbookpart.Workbook.Save();
    }

    static string NormSheetName(string? value, uint sheetId, HashSet<string> existNames)
    {
        value = RegularExpressions.InvalidSheetNameChars().Replace(value ?? string.Empty, string.Empty).Trim();

        if (string.IsNullOrWhiteSpace(value) || existNames.Contains(value))
            return $"Sheet{sheetId}";

        if (value.Length > _maxSheetName)
            return value.Substring(0, _maxSheetName);

        return value;
    }

    static void AddStyles(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet();

        // fonts
        stylesPart.Stylesheet.Fonts = new Fonts();
        stylesPart.Stylesheet.Fonts.AppendChild(new Font());
        stylesPart.Stylesheet.Fonts.AppendChild(new Font(new Bold(), new Color() { Rgb = HexBinaryValue.FromString("FFFFFFFF") }));
        stylesPart.Stylesheet.Fonts.AppendChild(new Font(new Underline(), new Color() { Theme = 10U }));
        stylesPart.Stylesheet.Fonts.Count = (uint)stylesPart.Stylesheet.Fonts.ChildElements.Count;

        // fills
        stylesPart.Stylesheet.Fills = new Fills();
        stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
        stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
        stylesPart.Stylesheet.Fills.AppendChild(new Fill
        {
            PatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF4F81BD") },
                BackgroundColor = new BackgroundColor { Indexed = 64 },
            },
        });
        stylesPart.Stylesheet.Fills.Count = (uint)stylesPart.Stylesheet.Fills.ChildElements.Count;

        // borders
        stylesPart.Stylesheet.Borders = new Borders();
        stylesPart.Stylesheet.Borders.AppendChild(new Border());
        stylesPart.Stylesheet.Borders.Count = (uint)stylesPart.Stylesheet.Borders.ChildElements.Count;

        // NumberingFormats
        //uint iExcelIndex = 164;
        stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
        stylesPart.Stylesheet.NumberingFormats.AddChild(new NumberingFormat { NumberFormatId = 3453, FormatCode = "0.00%" });
        stylesPart.Stylesheet.NumberingFormats.Count = (uint)stylesPart.Stylesheet.NumberingFormats.ChildElements.Count;

        // cell style formats
        stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
        stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());
        stylesPart.Stylesheet.CellStyleFormats.Count = 1;

        // cell styles
        stylesPart.Stylesheet.CellFormats = new CellFormats();
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
        // header style
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = false });
        // datetime style
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Top });
        // hyperlink style
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 2 }).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top });
        // multiline style
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = true });
        // percentage
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { NumberFormatId = 3453 }).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top });
        // nowrap style
        stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = false });

        stylesPart.Stylesheet.CellFormats.Count = (uint)stylesPart.Stylesheet.CellFormats.ChildElements.Count;

        stylesPart.Stylesheet.Save();
    }

    static IEnumerable<Row> GetRows(IEnumerable items, List<ColumnSchema> columns, bool wrapText)
    {
        var headerCells = columns.Select((x, i) => new Cell
        {
            CellReference = GetColReference(i),
            CellValue = new CellValue(x.Name),
            DataType = CellValues.String,
            StyleIndex = 1,
        }).ToArray();

        var headerRow = new Row() { RowIndex = 1 };
        headerRow.Append(headerCells);

        yield return headerRow;

        var i = 1u;
        foreach (var item in items)
        {
            var row = new Row() { RowIndex = ++i };
            row.Append(columns.Select((x, j) => GetCell(i, headerCells[j].CellReference, x.Value?.Invoke(item), wrapText)));
            yield return row;
        }
    }

    static Cell GetCell(uint rowIndex, string? cellReference, object? value, bool wrapText)
    {
        var cell = new Cell { CellReference = cellReference };

        if (value is ICellValue cellValue)
            cellValue.Apply(cell, rowIndex);
        else if (value is string str)
            CellText.Apply(cell, str, wrapText);
        else if (value is Uri uri)
            new CellHyperlink(uri).Apply(cell, rowIndex);
        else
            CellDefault.Apply(cell, value);

        return cell;
    }

    static string GetColReference(int index)
    {
        var result = new List<char>();
        while (index >= _digits.Length)
        {
            int remainder = index % _digits.Length;
            index = index / _digits.Length - 1;
            result.Add(_digits[remainder]);
        }
        result.Add(_digits[index]);
        result.Reverse();
        return new string([.. result]);
    }

    static readonly string _digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    const int _maxSheetName = 31;
}
