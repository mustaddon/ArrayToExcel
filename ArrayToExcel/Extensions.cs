using System;
using System.Collections.Generic;
using System.IO;

namespace ArrayToExcel;

public static partial class Extensions
{
    public static void ToExcel<T>(this IEnumerable<T> items, Stream stream, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(stream, items, schema);

    public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(items, schema).ToArray();

    public static MemoryStream ToExcelStream<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        => ExcelBuilder.Build(items, schema);

}
