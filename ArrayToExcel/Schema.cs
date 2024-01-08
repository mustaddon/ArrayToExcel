using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;

namespace ArrayToExcel
{
    internal class SheetSchema(string sheetName, List<ColumnSchema> columns, IEnumerable items)
    {
        public string SheetName { get; set; } = sheetName;
        public List<ColumnSchema> Columns { get; set; } = columns;
        public IEnumerable Items { get; set; } = items;
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
