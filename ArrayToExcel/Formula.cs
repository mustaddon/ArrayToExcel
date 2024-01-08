using System;

namespace ArrayToExcel
{
    public class Formula(Func<uint, string> rowText)
    {
        public Formula(string text) : this(row => text) { }

        internal Func<uint, string> RowText { get; } = rowText;
    }
}
