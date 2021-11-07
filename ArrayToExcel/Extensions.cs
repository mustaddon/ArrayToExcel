using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ArrayToExcel
{
    public static class Extensions
    {
        public static byte[] ToExcel<T>(this IEnumerable<T> items, Action<SchemaBuilder<T>>? schema = null)
        {
            return ArrayToExcel.CreateExcel(items, schema);
        }

        public static byte[] ToExcel(this DataSet dataSet, Action<SchemaBuilder<DataRow>>? schema = null)
        {
            var tables = dataSet.Tables.AsEnumerable().ToList();
            return ToExcel(tables.First(), builder =>
            {
                foreach (var table in tables.Skip(1))
                    builder.AddSheet(table);

                schema?.Invoke(builder);
            });
        }

        public static byte[] ToExcel(this DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
        {
            return ArrayToExcel.CreateExcel(dataTable.Rows.AsEnumerable(), builder =>
            {
                if (!string.IsNullOrWhiteSpace(dataTable.TableName))
                    builder.SheetName(dataTable.TableName);

                foreach (DataColumn col in dataTable.Columns)
                    builder.AddColumn(col.ColumnName, x => x[col]);

                schema?.Invoke(builder);
            });
        }

        public static SchemaBuilder<T> AddSheet<T>(this SchemaBuilder<T> builder, DataTable dataTable, Action<SheetSchemaBuilder<DataRow>>? schema = null)
        {
            return builder.AddSheet(dataTable.Rows.AsEnumerable(), builder =>
            {
                if (!string.IsNullOrWhiteSpace(dataTable.TableName))
                    builder.SheetName(dataTable.TableName);

                foreach (DataColumn col in dataTable.Columns)
                    builder.AddColumn(col.ColumnName, x => x[col]);

                schema?.Invoke(builder);
            });
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
