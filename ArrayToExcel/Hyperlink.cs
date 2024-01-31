using System;

namespace ArrayToExcel;

public class Hyperlink(string link, string? text = null)
{
    public Hyperlink(Uri link, string? text = null)
        : this(link.ToString(), text ?? link.OriginalString) { }

    private readonly Lazy<string> _format = new(() => Format(link, text));

    public string Link { get; } = link;
    public string? Text { get; } = text;

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
