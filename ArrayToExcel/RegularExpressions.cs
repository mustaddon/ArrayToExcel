using System.Text.RegularExpressions;

namespace ArrayToExcel
{
    internal static partial class RegularExpressions
    {
#if NET7_0_OR_GREATER
        [GeneratedRegex(@"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled)]
        public static partial Regex InvalidXmlChars();

        [GeneratedRegex(@"[:?*\\/\[\]\r\n]|[\uDC00-\uDFFF]|[\uD800-\uDBFF]|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled)]
        public static partial Regex InvalidSheetNameChars();
#else
        static readonly Regex _invalidXmlChars = new(@"(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled);
        public static Regex InvalidXmlChars() => _invalidXmlChars;

        static readonly Regex _invalidSheetNameChars = new(@"[:?*\\/\[\]\r\n]|[\uDC00-\uDFFF]|[\uD800-\uDBFF]|[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\uFFFE\uFFFF]", RegexOptions.Compiled);
        public static Regex InvalidSheetNameChars() => _invalidSheetNameChars; 
#endif
    }
}
