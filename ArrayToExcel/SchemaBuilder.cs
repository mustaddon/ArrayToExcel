using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
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

        public SchemaBuilder<T> ColumnName(Func<MemberInfo, string> name)
        {
            _rootSheetBuilder.ColumnName(name);
            return this;
        }

        public SchemaBuilder<T> ColumnWidth(Func<MemberInfo, uint> width)
        {
            _rootSheetBuilder.ColumnWidth(width);
            return this;
        }

        public SchemaBuilder<T> ColumnFilter(Func<MemberInfo, bool> filter)
        {
            _rootSheetBuilder.ColumnFilter(filter);
            return this;
        }

        public SchemaBuilder<T> ColumnSort<TKey>(Func<MemberInfo, TKey> sort, bool desc = false)
        {
            _rootSheetBuilder.ColumnSort(sort, desc);
            return this;
        }

        public SchemaBuilder<T> ColumnValue(Func<MemberInfo, T, object?> value)
        {
            _rootSheetBuilder.ColumnValue(value);
            return this;
        }

        public SchemaBuilder<T> AddColumn(string name, Func<T, object?> value, uint width = 20)
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

        public SheetSchemaBuilder<T> ColumnName(Func<MemberInfo, string> name)
        {
            foreach (var col in Schema.Columns)
                if (col.Member != null)
                    col.Name = name(col.Member);
            return this;
        }

        public SheetSchemaBuilder<T> ColumnWidth(Func<MemberInfo, uint> width)
        {
            foreach (var col in Schema.Columns)
                if (col.Member != null)
                    col.Width = width(col.Member);
            return this;
        }

        public SheetSchemaBuilder<T> ColumnFilter(Func<MemberInfo, bool> filter)
        {
            Schema.Columns = Schema.Columns.Where(x => x.Member == null || filter(x.Member)).ToList();
            return this;
        }

        public SheetSchemaBuilder<T> ColumnSort<TKey>(Func<MemberInfo, TKey> sort, bool desc = false)
        {
            if (!_defaultCols)
                return this;

            Schema.Columns = (desc
                ? Schema.Columns.OrderByDescending(x => x.Member != null ? sort(x.Member) : default)
                : Schema.Columns.OrderBy(x => x.Member != null ? sort(x.Member) : default)
            ).ToList();

            return this;
        }

        public SheetSchemaBuilder<T> ColumnValue(Func<MemberInfo, T, object?> value)
        {
            foreach (var col in Schema.Columns)
                if (col.Member != null)
                    col.Value = x => value(col.Member, (T)x);
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
