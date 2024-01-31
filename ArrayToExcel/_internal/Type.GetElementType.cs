using System;
using System.Collections;
using System.Collections.Generic;

namespace ArrayToExcel._internal;

internal static partial class TypeExt
{
    internal static Type? GetElementTypeExt(this Type type)
    {
        if (type.IsArray)
            return type.GetElementType();

        if (type.IsInterface && type.IsGenericType && type.GetGenericTypeDefinition() == _iEnumerable1)
            return type.GenericTypeArguments[0];

        var iEnumerable1 = type.GetInterface(_iEnumerable1.Name);

        if (iEnumerable1 != null)
            return iEnumerable1.GenericTypeArguments[0];

        if (typeof(IEnumerable).IsAssignableFrom(type))
            return typeof(object);

        return null;
    }

    static readonly Type _iEnumerable1 = typeof(IEnumerable<>);
}

