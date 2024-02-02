using System.Collections;
using System.Collections.Generic;

namespace ArrayToExcel._internal;

internal class SheetSchema(string sheetName, List<ColumnSchema> columns, IEnumerable items)
{
    public string SheetName = sheetName;
    public List<ColumnSchema> Columns = columns;
    public IEnumerable Items = items;
    public bool? WrapText;
    public bool? DateOnly;
}

