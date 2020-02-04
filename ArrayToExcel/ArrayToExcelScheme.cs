using System;
using System.Collections.Generic;
using System.Text;

namespace RandomSolutions
{
    public class ArrayToExcelScheme<T>
    {
        internal ArrayToExcelScheme() { }

        internal Dictionary<string, Func<T, object>> Columns = new Dictionary<string, Func<T, object>>();

        public ArrayToExcelScheme<T> AddColumn(string name, Func<T,object> value)
        {
            if (Columns.ContainsKey(name))
                Columns[name] = value;
            else
                Columns.Add(name, value);

            return this;
        }
    }
}
