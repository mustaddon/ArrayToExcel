using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace ArrayToExcel
{
    public static partial class Extensions
    {
        public static void ToExcel(this DataTable dataTable, Stream stream, Action<SchemaBuilder<DataRow>>? schema = null)
            => ArrayToExcel.CreateExcel(stream, dataTable.Rows.AsEnumerable(), b => dataTable.SchemaBuilder(b, schema));

        public static byte[] ToExcel(this DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
            => ToExcelStream(dataTable, schema).ToArray();

        public static MemoryStream ToExcelStream(this DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
            => ArrayToExcel.CreateExcel(dataTable.Rows.AsEnumerable(), b => dataTable.SchemaBuilder(b, schema));

        public static SchemaBuilder<T> AddSheet<T>(this SchemaBuilder<T> builder, DataTable dataTable, Action<SchemaBuilder<DataRow>>? schema = null)
            => builder.AddSheet(dataTable.Rows.AsEnumerable(), b => dataTable.SchemaBuilder(b, schema));


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
