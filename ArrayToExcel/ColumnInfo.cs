using System.Reflection;

namespace ArrayToExcel
{
    public sealed class ColumnInfo
    {
        internal ColumnInfo(int index, ColumnSchema schema)
        {
            Index = index;
            Schema = schema;
        }

        internal ColumnSchema Schema { get; }

        public int Index { get; }
        public string Name { get => this.Schema.Name; }
        public MemberInfo? Member { get => this.Schema.Member; }
    }
}
