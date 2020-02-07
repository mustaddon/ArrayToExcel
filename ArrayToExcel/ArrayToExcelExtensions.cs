using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace RandomSolutions
{
    public static class ArrayToExcelExtensions
    {
        public static byte[] ToExcel<T>(this IEnumerable<T> items, string sheetName = null)
        {
            return ArrayToExcel.CreateExcel(items, scheme =>
            {
                scheme.SheetName = sheetName;
            });
        }

        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<ArrayToExcelScheme<T>> schemeBuilder)
        {
            return ArrayToExcel.CreateExcel(items, schemeBuilder);
        }

    }
}
