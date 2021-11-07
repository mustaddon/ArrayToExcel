using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ArrayToExcel
{
    public class SchemaBuilder<T>
    {
        public SchemaBuilder(IEnumerable items)
        {
            _rootSheetBuilder = new(DefaultSheetName(), items);
            SheetSchemas.Add(_rootSheetBuilder.Schema);
        }

        private readonly SheetSchemaBuilder<T> _rootSheetBuilder;

        internal List<SheetSchema> SheetSchemas { get; } = new();

        private string DefaultSheetName() => $"Sheet{SheetSchemas.Count + 1}";

        public SchemaBuilder<T> AddSheet<TList>(IEnumerable<TList> list, Action<SheetSchemaBuilder<TList>>? schema = null)
        {
            var builder = new SheetSchemaBuilder<TList>(DefaultSheetName(), list);
            schema?.Invoke(builder);
            SheetSchemas.Add(builder.Schema);
            return this;
        }

        public SchemaBuilder<T> SheetName(string name)
        {
            _rootSheetBuilder.SheetName(name);
            return this;
        }

        public SchemaBuilder<T> ColumnName(Func<ColumnInfo, string> name)
        {
            _rootSheetBuilder.ColumnName(name);
            return this;
        }

        public SchemaBuilder<T> ColumnWidth(Func<ColumnInfo, uint> width)
        {
            _rootSheetBuilder.ColumnWidth(width);
            return this;
        }

        public SchemaBuilder<T> ColumnFilter(Func<ColumnInfo, bool> filter)
        {
            _rootSheetBuilder.ColumnFilter(filter);
            return this;
        }

        public SchemaBuilder<T> ColumnSort<TKey>(Func<ColumnInfo, TKey> sort, bool desc = false)
        {
            _rootSheetBuilder.ColumnSort(sort, desc);
            return this;
        }

        public SchemaBuilder<T> ColumnValue(Func<ColumnInfo, T, object?> value)
        {
            _rootSheetBuilder.ColumnValue(value);
            return this;
        }

        public SchemaBuilder<T> AddColumn(string name, Func<T, object?> value, uint width = ColumnSchema.DefaultWidth)
        {
            _rootSheetBuilder.AddColumn(name, value, width);
            return this;
        }
    }

    public class SheetSchemaBuilder<T>
    {
        public SheetSchemaBuilder(string sheetName, IEnumerable items)
        {
            Schema = new(sheetName, DefaultColumns(items), items);
        }

        private bool _defaultCols = true;

        internal SheetSchema Schema { get; }

        public SheetSchemaBuilder<T> SheetName(string name)
        {
            Schema.SheetName = name;
            return this;
        }

        public SheetSchemaBuilder<T> ColumnName(Func<ColumnInfo, string> name)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Name = name(col);
            return this;
        }

        public SheetSchemaBuilder<T> ColumnWidth(Func<ColumnInfo, uint> width)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Width = width(col);
            return this;
        }

        public SheetSchemaBuilder<T> ColumnFilter(Func<ColumnInfo, bool> filter)
        {
            Schema.Columns = Schema.Columns
                .Select((x, i) => new ColumnInfo(i, x))
                .Where(x => filter(x))
                .Select(x => x.Schema)
                .ToList();
            return this;
        }

        public SheetSchemaBuilder<T> ColumnSort<TKey>(Func<ColumnInfo, TKey> sort, bool desc = false)
        {
            var colInfos = Schema.Columns.Select((x, i) => new ColumnInfo(i, x)).ToList();

            Schema.Columns = (desc
                ? colInfos.OrderByDescending(sort)
                : colInfos.OrderBy(sort)
            ).Select(x => x.Schema).ToList();

            return this;
        }

        public SheetSchemaBuilder<T> ColumnValue(Func<ColumnInfo, T, object?> value)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Value = x => value(col, (T)x);
            return this;
        }

        public SheetSchemaBuilder<T> AddColumn(string name, Func<T, object?> value, uint width = ColumnSchema.DefaultWidth)
        {
            if (_defaultCols)
            {
                Schema.Columns.Clear();
                _defaultCols = false;
            }

            Schema.Columns.Add(new()
            {
                Name = name,
                Value = x => value((T)x),
                Width = width,
            });

            return this;
        }

        private List<ColumnSchema> DefaultColumns(IEnumerable items)
        {
            var type = typeof(T);

            if (typeof(IDictionary<string, object?>).IsAssignableFrom(type))
            {
                var enumerator = items.GetEnumerator();
                enumerator.MoveNext();
                return (enumerator.Current as IDictionary<string, object?>)
                    ?.Select(kvp => new ColumnSchema()
                    {
                        Name = kvp.Key,
                        Value = new(x => (x as IDictionary<string, object?>)?[kvp.Key]),
                    })
                    .ToList() ?? new List<ColumnSchema>();
            }

            if (typeof(IDictionary).IsAssignableFrom(type))
            {
                var enumerator = items.GetEnumerator();
                enumerator.MoveNext();
                var dict = (enumerator.Current as IDictionary)?.GetEnumerator();

                var result = new List<ColumnSchema>();

                while (dict?.MoveNext() == true)
                {
                    var key = dict.Key;
                    result.Add(new()
                    {
                        Name = key.ToString(),
                        Value = new(x => (x as IDictionary)?[key]),
                    });
                }

                return result;
            }

            return type.GetMembers(BindingFlags.Instance | BindingFlags.Public)
                .Where(x => x is PropertyInfo || x is FieldInfo)
                .Select(member => new ColumnSchema
                {
                    Member = member,
                    Name = member.Name,
                    Value = new(x => (member as PropertyInfo)?.GetValue(x) ?? (member as FieldInfo)?.GetValue(x)),
                })
                .ToList();
        }

    }
}
