using System;
using System.Collections.Generic;
using System.IO;

namespace ArrayToExcel
{
    public static partial class Extensions
    {
        public static void ToExcel<T>(this IEnumerable<T> items, Stream stream, Action<SchemaBuilder<T>>? schema = null)
            => ArrayToExcel.CreateExcel(stream, items, schema);

        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
            => ArrayToExcel.CreateExcel(items, schema).ToArray();

        public static MemoryStream ToExcelStream<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
            => ArrayToExcel.CreateExcel(items, schema);

    }
}
