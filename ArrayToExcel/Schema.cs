using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace ArrayToExcel
{
    internal class SheetSchema
    {
        public SheetSchema(string sheetName, List<ColumnSchema> columns, IEnumerable items)
        {
            SheetName = sheetName;
            Columns = columns;
            Items = items;
        }

        public string SheetName { get; set; }
        public List<ColumnSchema> Columns { get; set; }
        public IEnumerable Items { get; set; }
    }

    internal class ColumnSchema
    {
        public MemberInfo? Member { get; set; }
        public uint Width { get; set; } = DefaultWidth;
        public string Name { get; set; } = string.Empty;
        public Func<object, object?>? Value { get; set; }

        public const uint DefaultWidth = 20;
    }
}
