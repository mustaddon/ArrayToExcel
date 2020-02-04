using System;
using System.Collections.Generic;
using System.Text;

namespace RandomSolutions
{
    public static class ArrayToExcelExtensions
    {
        public static byte[] ToExcel<T>(this IEnumerable<T> items)
            => ArrayToExcel.CreateExcel(items);

        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<ArrayToExcelScheme<T>> schemeBuilder)
            => ArrayToExcel.CreateExcel(items, schemeBuilder);

        public static byte[] ToExcel<T>(this IEnumerable<T> items, Dictionary<string, Func<T, object>> columns)
            => ArrayToExcel.CreateExcel(items, columns);
    }
}
