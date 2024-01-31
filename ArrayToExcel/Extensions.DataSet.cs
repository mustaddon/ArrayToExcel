using System;
using System.Data;
using System.IO;
using System.Linq;

namespace ArrayToExcel;

public static partial class Extensions
{
    public static void ToExcel(this DataSet dataSet, Stream stream, Action<DataTable, SchemaBuilder<DataRow>>? schema = null)
    {
        var tables = dataSet.Tables.AsEnumerable().ToList();
        ToExcel(tables.First(), stream, builder =>
        {
            foreach (var table in tables.Skip(1))
                builder.AddSheet(table, s => schema?.Invoke(table, s));

            schema?.Invoke(tables.First(), builder);
        });
    }

    public static byte[] ToExcel(this DataSet dataSet, Action<DataTable, SchemaBuilder<DataRow>>? schema = null)
        => ToExcelStream(dataSet, schema).ToArray();

    public static MemoryStream ToExcelStream(this DataSet dataSet, Action<DataTable, SchemaBuilder<DataRow>>? schema = null)
    {
        var ms = new MemoryStream();
        ToExcel(dataSet, ms, schema);
        ms.Position = 0;
        return ms;
    }
}
