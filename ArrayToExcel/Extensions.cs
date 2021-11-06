using System;
using System.Collections.Generic;

namespace ArrayToExcel
{
    public static class Extensions
    {
        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            return ArrayToExcel.CreateExcel(items, schema);
        }

    }
}
