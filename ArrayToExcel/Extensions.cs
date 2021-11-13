using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ArrayToExcel
{
    public static class Extensions
    {
        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            using var ms = ToExcelStream(items, schema);
            return ms.ToArray();
        }

        public static byte[] ToExcel(this DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
        {
            using var ms = ToExcelStream(dataTable, schema);
            return ms.ToArray();
        }

        public static byte[] ToExcel(this DataSet dataSet, Action<DataTable, SchemaBuilder<DataRow>>? schema = null)
        {
            using var ms = ToExcelStream(dataSet, schema);
            return ms.ToArray();
        }

        public static MemoryStream ToExcelStream<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            return ArrayToExcel.CreateExcel(items, schema);
        }

        public static MemoryStream ToExcelStream(this DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
        {
            return ArrayToExcel.CreateExcel(dataTable.Rows.AsEnumerable(), b => dataTable.SchemaBuilder(b, schema));
        }

        public static MemoryStream ToExcelStream(this DataSet dataSet, Action<DataTable, SchemaBuilder<DataRow>>? schema = null)
        {
            var tables = dataSet.Tables.AsEnumerable().ToList();
            return ToExcelStream(tables.First(), builder =>
            {
                foreach (var table in tables.Skip(1))
                    builder.AddSheet(table, s => schema?.Invoke(table, s));

                schema?.Invoke(tables.First(), builder);
            });
        }

        public static SchemaBuilder<T> AddSheet<T>(this SchemaBuilder<T> builder, DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
        {
            return builder.AddSheet(dataTable.Rows.AsEnumerable(), b => dataTable.SchemaBuilder(b, schema));
        }


        private static void SchemaBuilder(this DataTable dataTable, SchemaBuilder<DataRow> builder, Action<SchemaBuilder<DataRow>>? schema)
        {
            if (!string.IsNullOrWhiteSpace(dataTable.TableName))
                builder.SheetName(dataTable.TableName);

            foreach (DataColumn col in dataTable.Columns)
                builder.AddColumn(col.ColumnName, x => x[col]);

            schema?.Invoke(builder);
        }

        private static IEnumerable<DataRow> AsEnumerable(this DataRowCollection items)
        {
            foreach (DataRow item in items)
                yield return item;
        }

        private static IEnumerable<DataTable> AsEnumerable(this DataTableCollection items)
        {
            foreach (DataTable item in items)
                yield return item;
        }
    }
}
