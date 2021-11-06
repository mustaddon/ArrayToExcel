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
        public static byte[] CreateExcel<T>(IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            var builder = new SchemaBuilder<T>(items);
            schema?.Invoke(builder);
            return _createExcel(builder.SheetSchemas);
        }

        static byte[] _createExcel(List<SheetSchema> sheetSchemas)
        {
            using (var ms = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                    _addStyles(document);

                    var sheetId = 1u;

                    foreach (var sheetSchema in sheetSchemas)
                    {
                        var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet();

                        sheets.Append(new Sheet()
                        {
                            Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = sheetId++,
                            Name = _normSheetName(sheetSchema.SheetName),
                        });

                        if (sheetSchema.Columns.Count > 0)
                        {
                            var cols = worksheetPart.Worksheet.AppendChild(new Columns());
                            cols.Append(sheetSchema.Columns.Select((x, i) => new Column() { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = x.Width, CustomWidth = true, BestFit = true }));

                            var rows = _getRows(sheetSchema.Items, sheetSchema.Columns);
                            var sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());
                            sheetData.Append(rows);

                            worksheetPart.Worksheet.Append(new AutoFilter() { Reference = $"A1:{_getColReference(sheetSchema.Columns.Count - 1)}{rows.Length}" });
                        }
                    }

                    workbookpart.Workbook.Save();
                }
                return ms.ToArray();
            }
        }

        static string _normSheetName(string value)
        {
            return value.Length > 31 ? value.Substring(0, 28) + "..." : value;
        }

        static void _addStyles(SpreadsheetDocument document)
        {
            var stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            // fonts
            stylesPart.Stylesheet.Fonts = new Fonts();
            stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            var font1 = stylesPart.Stylesheet.Fonts.AppendChild(new Font());
            font1.Append(new Bold());
            font1.Append(new Color() { Rgb = HexBinaryValue.FromString("FFFFFFFF") });
            stylesPart.Stylesheet.Fonts.Count = (uint)stylesPart.Stylesheet.Fonts.ChildElements.Count;

            // fills
            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            var fill2 = stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill() { PatternType = PatternValues.Solid } });
            fill2.PatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF4F81BD") };
            fill2.PatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
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
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center });
            // datetime style
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { ApplyNumberFormat = true, NumberFormatId = 14, FormatId = 0, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = true });
            stylesPart.Stylesheet.CellFormats.Count = (uint)stylesPart.Stylesheet.CellFormats.ChildElements.Count;

            stylesPart.Stylesheet.Save();
        }

        static Row[] _getRows(IEnumerable items, List<ColumnSchema> columns)
        {
            var rows = new List<Row>();

            var headerCells = columns.Select((x, i) => new Cell
            {
                CellReference = _getColReference(i),
                CellValue = new CellValue(x.Name),
                DataType = CellValues.String,
                StyleIndex = 1,
            }).ToArray();

            var headerRow = new Row() { RowIndex = 1 };
            headerRow.Append(headerCells);
            rows.Add(headerRow);

            var i = 2;
            foreach (var item in items)
            {
                var row = new Row() { RowIndex = (uint)i++ };
                var cells = columns.Select((x, i) => _getCell(headerCells[i].CellReference, x.Value?.Invoke(item))).ToArray();
                row.Append(cells);
                rows.Add(row);
            }
            return rows.ToArray();
        }

        static Cell _getCell(string reference, object? value)
        {
            var dataType = _getCellType(value);
            return new Cell
            {
                CellReference = reference,
                CellValue = _getCellValue(value),
                DataType = dataType,
                StyleIndex = dataType == CellValues.Date ? 2 : 0u,
            };
        }

        static CellValue _getCellValue(object? value)
        {
            if (value == null) return new CellValue();

            var type = value.GetType();

            if (type == typeof(bool))
                return new CellValue((bool)value ? "1" : "0");

            if (type == typeof(DateTime))
                return new CellValue(((DateTime)value).ToString("s", _cultureInfo));

            if (type == typeof(DateTimeOffset))
                return new CellValue(((DateTimeOffset)value).ToString("s", _cultureInfo));

            if (type == typeof(double))
                return new CellValue(((double)value).ToString(_cultureInfo));

            if (type == typeof(decimal))
                return new CellValue(((decimal)value).ToString(_cultureInfo));

            if (type == typeof(float))
                return new CellValue(((float)value).ToString(_cultureInfo));

            return new CellValue(_removeInvalidXmlChars(value.ToString()));
        }

        static string _removeInvalidXmlChars(string text)
        {
            return string.IsNullOrEmpty(text) ? text
                : new Regex(@"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled)
                    .Replace(text, string.Empty);
        }

        static CellValues _getCellType(object? value)
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

        static string _getColReference(int index)
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

        static HashSet<Type> _numericTypes = new HashSet<Type>
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

        static CultureInfo _cultureInfo = CultureInfo.GetCultureInfo("en-US");

        static string _digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    }
}
