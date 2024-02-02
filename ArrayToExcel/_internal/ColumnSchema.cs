using System;
using System.Reflection;

namespace ArrayToExcel._internal;

internal class ColumnSchema
{
    public MemberInfo? Member;
    public uint Width = DefaultWidth;
    public string Name = string.Empty;
    public Func<object, object?>? Value;

    public const uint DefaultWidth = 20;
}

