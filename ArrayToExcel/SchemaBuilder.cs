using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ArrayToExcel
{
    public sealed class SchemaBuilder<T>
    {
        internal SchemaBuilder(IEnumerable items, List<SheetSchema>? parentChilds = null)
        {
            Childs = parentChilds ?? new();
            Schema = new($"Sheet{(parentChilds?.Count + 2) ?? 1}", DefaultColumns(items), items);
        }

        private bool _defaultCols = true;

        internal SheetSchema Schema { get; }
        internal List<SheetSchema> Childs { get; }

        public SchemaBuilder<T> SheetName(string name)
        {
            Schema.SheetName = name;
            return this;
        }

        public SchemaBuilder<T> ColumnName(Func<ColumnInfo, string> name)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Name = name(col);
            return this;
        }

        public SchemaBuilder<T> ColumnWidth(Func<ColumnInfo, uint> width)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Width = width(col);
            return this;
        }

        public SchemaBuilder<T> ColumnFilter(Func<ColumnInfo, bool> filter)
        {
            Schema.Columns = Schema.Columns
                .Select((x, i) => new ColumnInfo(i, x))
                .Where(x => filter(x))
                .Select(x => x.Schema)
                .ToList();
            return this;
        }

        public SchemaBuilder<T> ColumnSort<TKey>(Func<ColumnInfo, TKey> sort, bool desc = false)
        {
            var colInfos = Schema.Columns.Select((x, i) => new ColumnInfo(i, x)).ToList();

            Schema.Columns = (desc
                ? colInfos.OrderByDescending(sort)
                : colInfos.OrderBy(sort)
            ).Select(x => x.Schema).ToList();

            return this;
        }

        public SchemaBuilder<T> ColumnValue(Func<ColumnInfo, T, object?> value)
        {
            foreach (var col in Schema.Columns.Select((x, i) => new ColumnInfo(i, x)))
                col.Schema.Value = x => value(col, (T)x);
            return this;
        }

        public SchemaBuilder<T> AddColumn(string name, Func<T, object?> value, uint width = ColumnSchema.DefaultWidth)
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

        public SchemaBuilder<T> AddSheet<TList>(IEnumerable<TList> list, Action<SchemaBuilder<TList>>? schema = null)
        {
            var builder = new SchemaBuilder<TList>(list, Childs);
            schema?.Invoke(builder);
            Childs.Add(builder.Schema);
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
