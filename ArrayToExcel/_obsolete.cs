using System;
using System.Collections.Generic;
using System.IO;

namespace ArrayToExcel;

[Obsolete("Class is deprecated, please use ExcelBuilder instead.")]
public class ArrayToExcel
{
    [Obsolete("Method is deprecated, please use ExcelBuilder.Build instead.")]
    public static MemoryStream CreateExcel<T>(IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(items, schema);

    [Obsolete("Method is deprecated, please use ExcelBuilder.Build instead.")]
    public static void CreateExcel<T>(Stream stream, IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(stream, items, schema);
}
