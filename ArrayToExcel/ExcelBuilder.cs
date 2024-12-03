using ArrayToExcel._internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
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
    public static bool DefaultDateOnly = false;

    static void CreateExcel(Stream stream, IEnumerable<SheetSchema> sheetSchemas)
    {
        using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

        var workbookpart = document.AddWorkbookPart();
        var workbook = workbookpart.Workbook = new();
        var sheets = workbook.Sheets = new();
        var definedNames = workbook.DefinedNames = new();

        AddStyles(workbookpart);

        var sheetNames = new HashSet<string>();
        var sheetId = 0u;

        foreach (var sheetSchema in sheetSchemas)
        {
            sheetId++;

            var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            var worksheet = worksheetPart.Worksheet = new();
            var dimension = worksheet.SheetDimension = new() { Reference = "A1" };
            var cols = worksheet.AppendChild(new Columns());
            var sheetData = worksheet.AppendChild(new SheetData());

            var sheet = new Sheet()
            {
                Id = workbookpart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = NormSheetName(sheetSchema.SheetName, sheetId, sheetNames),
            };

            sheetNames.Add(sheet.Name.Value ?? string.Empty);
            sheets.AppendChild(sheet);

            if (sheetSchema.Columns.Count == 0)
            {
                cols.AppendChild(new Column() { Min = 1, Max = 1, BestFit = true });
                continue;
            }

            cols.Append(sheetSchema.Columns.Select((x, i) => new Column() { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = x.Width, CustomWidth = true, BestFit = true }));

            sheetData.Append(GetRows(sheetSchema));

            dimension.Reference = $"A1:{GetColReference(sheetSchema.Columns.Count - 1)}{sheetData.ChildElements.Count}";
            worksheet.AppendChild(new AutoFilter() { Reference = $"A1:{GetColReference(sheetSchema.Columns.Count - 1)}1" });
            definedNames.AppendChild(new DefinedName($"'{sheet.Name.Value!.Replace("'", "''")}'!$A$1:${GetColReference(sheetSchema.Columns.Count - 1)}$1") { Name = "_xlnm._FilterDatabase", LocalSheetId = sheetId - 1, Hidden = true });
        }

        workbook.Save();
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
        var stylesheet = stylesPart.Stylesheet = new();

        // fonts
        var fonts = stylesheet.Fonts = new();
        fonts.AppendChild(new Font());
        fonts.AppendChild(new Font(new Bold(), new Color() { Rgb = HexBinaryValue.FromString("FFFFFFFF") }));
        fonts.AppendChild(new Font(new Underline(), new Color() { Theme = 10U }));
        fonts.Count = (uint)fonts.ChildElements.Count;

        // fills
        var fills = stylesheet.Fills = new();
        fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
        fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
        fills.AppendChild(new Fill
        {
            PatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF4F81BD") },
                BackgroundColor = new BackgroundColor { Indexed = 64 },
            },
        });
        fills.Count = (uint)fills.ChildElements.Count;

        // borders
        var borders = stylesheet.Borders = new();
        borders.AppendChild(new Border());
        borders.Count = (uint)borders.ChildElements.Count;

        // NumberingFormats
        //uint iExcelIndex = 164;
        var numberingFormats = stylesheet.NumberingFormats = new NumberingFormats();
        numberingFormats.AddChild(new NumberingFormat { NumberFormatId = 3453, FormatCode = "0.00%" });
        numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;

        // cell style formats
        stylesheet.CellStyleFormats = new();
        stylesheet.CellStyleFormats.AppendChild(new CellFormat());
        stylesheet.CellStyleFormats.Count = 1;

        // cell styles
        var cellFormats = stylesheet.CellFormats = new();
        cellFormats.AppendChild(new CellFormat());
        // header style
        cellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = false });
        // default style
        cellFormats.AppendChild(new CellFormat()).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = false });
        // wraptext style
        cellFormats.AppendChild(new CellFormat()).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = true });
        // date style
        cellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Top });
        // datetime style
        cellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 22, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Top });
        // hyperlink style
        cellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 2 }).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top });
        // percentage style
        cellFormats.AppendChild(new CellFormat { NumberFormatId = 3453 }).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top });

        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        //stylesheet.Save();
    }

    static IEnumerable<Row> GetRows(SheetSchema sheetSchema)
    {
        var settings = new SheetSettings
        {
            WrapText = sheetSchema.WrapText ?? DefaultWrapText,
            DateOnly = sheetSchema.DateOnly ?? DefaultDateOnly,
        };
        var columns = sheetSchema.Columns;

        var headerRow = new Row() { RowIndex = 1 };

        headerRow.Append(columns.Select((x, i) => new Cell
        {
            CellReference = GetColReference(i) + "1",
            CellValue = new CellValue(x.Name),
            DataType = CellValues.String,
            StyleIndex = Styles.Header,
        }));

        yield return headerRow;

        var i = 1u;
        foreach (var item in sheetSchema.Items)
        {
            var row = new Row() { RowIndex = ++i };
            row.Append(columns.Select((x, j) => GetCell(i, GetColReference(j) + i, x.Value?.Invoke(item), settings)));
            yield return row;
        }
    }

    static Cell GetCell(uint rowIndex, string? cellReference, object? value, SheetSettings settings)
    {
        var cell = new Cell { CellReference = cellReference };

        if (value is ICellValue cellValue)
            cellValue.Apply(cell, rowIndex);
        else if (value is string str)
            CellText.Apply(cell, str, settings.WrapText);
        else if (value is DateTime dateTime)
            CellDate.Apply(cell, dateTime, settings.DateOnly);
        else if (value is DateTimeOffset dateTimeOffset)
            CellDate.Apply(cell, dateTimeOffset, settings.DateOnly);
#if NET6_0_OR_GREATER
        else if (value is DateOnly dateOnly)
            CellDate.Apply(cell, dateOnly);
#endif
        else if (value is Uri uri)
            CellHyperlink.Apply(cell, uri);
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
