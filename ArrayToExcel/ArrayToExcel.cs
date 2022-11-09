using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ArrayToExcel
{
    public class ArrayToExcel
    {
        public static MemoryStream CreateExcel<T>(IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            var builder = new SchemaBuilder<T>(items);
            schema?.Invoke(builder);
            return CreateExcel(new[] { builder.Schema }.Concat(builder.Childs));
        }

        static MemoryStream CreateExcel(IEnumerable<SheetSchema> sheetSchemas)
        {
            var ms = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
            {
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
                    sheets.Append(sheet);

                    var cols = worksheetPart.Worksheet.AppendChild(new Columns());
                    var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    if (sheetSchema.Columns.Count == 0)
                    {
                        cols.Append(new Column() { Min = 1, Max = 1, BestFit = true });
                        continue;
                    }

                    cols.Append(sheetSchema.Columns.Select((x, i) => new Column() { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = x.Width, CustomWidth = true, BestFit = true }));

                    sheetData.Append(GetRows(sheetSchema.Items, sheetSchema.Columns));

                    worksheetPart.Worksheet.Append(new AutoFilter() { Reference = $"A1:{GetColReference(sheetSchema.Columns.Count - 1)}{sheetData.ChildElements.Count}" });
                }

                workbookpart.Workbook.Save();
            }
            ms.Position = 0;
            return ms;
        }

        static string NormSheetName(string? value, uint sheetId, HashSet<string> existNames)
        {
            value = _invalidSheetNameChars.Replace(value ?? string.Empty, string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(value) || existNames.Contains(value))
                return $"Sheet{sheetId}";

            if (value.Length > 31)
                return $"{value.Substring(0, 28)}...";

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
            stylesPart.Stylesheet.NumberingFormats.Count = (uint)stylesPart.Stylesheet.NumberingFormats.ChildElements.Count;

            // cell style formats
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.Count = 1;

            // cell styles
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            // header style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center, WrapText = true });
            // datetime style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true }).AppendChild(new Alignment { Vertical = VerticalAlignmentValues.Top });
            // hyperlink style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 2 }).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top });
            // multiline style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat()).AppendChild(new Alignment() { Vertical = VerticalAlignmentValues.Top, WrapText = true });

            stylesPart.Stylesheet.CellFormats.Count = (uint)stylesPart.Stylesheet.CellFormats.ChildElements.Count;

            stylesPart.Stylesheet.Save();
        }

        static IEnumerable<Row> GetRows(IEnumerable items, List<ColumnSchema> columns)
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

            var i = 2u;
            foreach (var item in items)
            {
                var row = new Row() { RowIndex = i++ };
                var cells = columns.Select((x, i) => GetCell(row.RowIndex, headerCells[i].CellReference, x.Value?.Invoke(item))).ToArray();
                row.Append(cells);
                yield return row;
            }
        }

        static Cell GetCell(uint rowIndex, string? cellReference, object? value)
        {
            var cell = new Cell { CellReference = cellReference };

            if (value is string text)
            {
                cell.InlineString = GetInlineString(text);
                cell.DataType = CellValues.InlineString;
                cell.StyleIndex = 4;
            }
            else if (value is Formula formula)
            {
                cell.CellFormula = new CellFormula(formula.RowText(rowIndex));
                cell.StyleIndex = 4;
            }
            else if (value is Hyperlink hyperlink)
            {
                cell.CellFormula = new CellFormula(hyperlink.ToString());
                cell.StyleIndex = 3;
            }
            else if (value is Uri uri)
            {
                cell.CellFormula = new CellFormula(new Hyperlink(uri).ToString());
                cell.StyleIndex = 3;
            }
            else
            {
                cell.CellValue = GetCellValue(value);
                cell.DataType = GetCellType(value);
                cell.StyleIndex = cell.DataType == CellValues.Date ? 2 : 4u;
            }

            return cell;
        }

        static CellValue GetCellValue(object? value)
        {
            if (value == null) return new();

            var type = value.GetType();

            if (type == typeof(bool))
                return new((bool)value ? "1" : "0");

            if (type == typeof(DateTime))
                return new(((DateTime)value).ToString("s", _cultureInfo));

            if (type == typeof(DateTimeOffset))
                return new(((DateTimeOffset)value).ToString("s", _cultureInfo));

            if (type == typeof(double))
                return new(((double)value).ToString(_cultureInfo));

            if (type == typeof(decimal))
                return new(((decimal)value).ToString(_cultureInfo));

            if (type == typeof(float))
                return new(((float)value).ToString(_cultureInfo));

            return new(_invalidXmlChars.Replace(value.ToString()!, string.Empty));
        }

        static InlineString GetInlineString(string value)
        {
            return new InlineString(new Text(_invalidXmlChars.Replace(value, string.Empty))
            {
                Space = SpaceProcessingModeValues.Preserve
            });
        }

        static CellValues GetCellType(object? value)
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
            return new string(result.ToArray());
        }

        static readonly HashSet<Type> _numericTypes = new()
        {
            typeof(short),
            typeof(ushort),
            typeof(int),
            typeof(uint),
            typeof(long),
            typeof(ulong),
            typeof(double),
            typeof(decimal),
            typeof(float),
        };

        static readonly CultureInfo _cultureInfo = CultureInfo.GetCultureInfo("en-US");

        static readonly string _digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        static readonly Regex _invalidXmlChars = new(@"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled);

        static readonly Regex _invalidSheetNameChars = new(@"[:?*\\/\[\]\r\n]|[\uDC00-\uDFFF]|[\uD800-\uDBFF]|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled);
    }
}
