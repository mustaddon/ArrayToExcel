using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace RandomSolutions
{
    public class ArrayToExcel
    {
        public static byte[] CreateExcel<T>(IEnumerable<T> items)
        {
            return CreateExcel(items, typeof(T).GetMembers(BindingFlags.Instance | BindingFlags.Public)
                .Where(x => x is PropertyInfo || x is FieldInfo)
                .ToDictionary(m => m.Name, m => new Func<T, object>(x => (m as PropertyInfo)?.GetValue(x) ?? (m as FieldInfo)?.GetValue(x))));
        }

        public static byte[] CreateExcel<T>(IEnumerable<T> items, Action<ArrayToExcelScheme<T>> schemeBuilder)
        {
            var scheme = new ArrayToExcelScheme<T>();
            schemeBuilder.Invoke(scheme);
            return CreateExcel(items, scheme.Columns);
        }

        public static byte[] CreateExcel<T>(IEnumerable<T> items, Dictionary<string, Func<T, object>> columns)
        {
            using (var ms = new MemoryStream())
            {
                using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    var workbookpart = document.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();

                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
                    sheets.Append(new Sheet()
                    {
                        Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Sheet1"
                    });

                    _addStyles(document);

                    if (columns.Count > 0) {
                        var rows = _getRows(items, columns);
                        sheetData.Append(rows);
                        var range = $"A1:{_getColReference(columns.Count - 1)}{rows.Length}";
                        worksheetPart.Worksheet.Append(new AutoFilter() { Reference = range });
                    }

                    workbookpart.Workbook.Save();
                }
                return ms.ToArray();
            }
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
            stylesPart.Stylesheet.Fonts.Count = 2;

            // fills
            stylesPart.Stylesheet.Fills = new Fills();
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
            stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
            var fill2 = stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill() { PatternType = PatternValues.Solid } });
            fill2.PatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF4F81BD") };
            fill2.PatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            stylesPart.Stylesheet.Fills.Count = 3;

            // borders
            stylesPart.Stylesheet.Borders = new Borders();
            stylesPart.Stylesheet.Borders.AppendChild(new Border());
            stylesPart.Stylesheet.Borders.Count = 1;

            // cell style formats
            stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats();
            stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellStyleFormats.Count = 1;

            // cell styles
            stylesPart.Stylesheet.CellFormats = new CellFormats();
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
            stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 2, ApplyFill = true }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center });
            stylesPart.Stylesheet.CellFormats.Count = 2;

            stylesPart.Stylesheet.Save();
        }

        static Row[] _getRows<T>(IEnumerable<T> items, Dictionary<string, Func<T, object>> columns)
        {
            var rows = new List<Row>();
            var i = 0;
            var headerCells = columns.Select(x => new Cell
            {
                CellReference = _getColReference(i++),
                CellValue = new CellValue(x.Key),
                DataType = CellValues.String,
                StyleIndex = 1,
            }).ToArray();

            var headerRow = new Row() { RowIndex = 1 };
            headerRow.Append(headerCells);
            rows.Add(headerRow);

            i = 2;
            foreach (var item in items)
            {
                var row = new Row() { RowIndex = (uint)i++ };
                var j = 0;
                var cells = columns.Select(x =>
                {
                    var value = x.Value(item);
                    return new Cell
                    {
                        CellReference = headerCells[j++].CellReference,
                        CellValue = _getValue(value),
                        DataType = _getDataType(value),
                    };
                }).ToArray();
                row.Append(cells);
                rows.Add(row);
            }
            return rows.ToArray();
        }
        
        static string _getColReference(int data)
        {
            var result = new List<char>();
            while (data >= _digits.Length)
            {
                int remainder = data % _digits.Length;
                data = data /_digits.Length -1;
                result.Add(_digits[remainder]);
            }
            result.Add(_digits[data]);
            result.Reverse();
            return new string(result.ToArray());
        }
        
        static CellValue _getValue(object value)
        {
            if (value == null) return new CellValue();

            var type = value.GetType();

            if (type == typeof(DateTime))
                return new CellValue((DateTime)value);

            if (type == typeof(DateTimeOffset))
                return new CellValue((DateTimeOffset)value);

            return new CellValue(value.ToString());
        }

        static CellValues _getDataType(object value)
        {
            var type = value?.GetType();

            if (type == typeof(bool))
                return CellValues.Boolean;

            if (_numericTypes.Contains(type))
                return CellValues.Number;

            if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
                return CellValues.Date;

            return CellValues.String;
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

        static string _digits = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    }
}
