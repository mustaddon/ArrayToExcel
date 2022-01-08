using System;

namespace ArrayToExcel
{
    public class Hyperlink
    {
        public Hyperlink(Uri link, string? text = null)
            : this(link.ToString(), text ?? link.OriginalString) { }

        public Hyperlink(string link, string? text = null)
        {
            _format = new Lazy<string>(() => Format(link, text));

            Link = link;
            Text = text;
        }

        private readonly Lazy<string> _format;

        public string Link { get; }
        public string? Text { get; }

        public override string ToString() => _format.Value;


        private static string Format(string link, string? text)
        {
            return string.Format(string.IsNullOrWhiteSpace(text) || text == link
                ? "HYPERLINK(\"{0}\")"
                : "HYPERLINK(\"{0}\",\"{1}\")",
                Fix(link), Fix(text));
        }
        private static string? Fix(string? val) => val?.Replace("\"", "\"\"");

    }
}
