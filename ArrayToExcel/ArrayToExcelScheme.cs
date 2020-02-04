using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RandomSolutions
{
    public class ArrayToExcelScheme<T>
    {
        internal ArrayToExcelScheme(Dictionary<string, Func<T, object>> columns = null)
        {
            if(columns != null)
                foreach(var col in columns)
                    AddColumn(col.Key, col.Value);
        }

        internal List<Column> Columns = new List<Column>();

        public ArrayToExcelScheme<T> AddColumn(string name, Func<T, object> value, uint width = 20)
        {
            Columns.Add(new Column
            {
                Index = Columns.Count,
                Name = name,
                ValueFn = value,
                Width = width,
            });
            return this;
        }

        internal class Column
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public Func<T, object> ValueFn { get; set; }
            public uint Width { get; set; }
        }
    }
}
