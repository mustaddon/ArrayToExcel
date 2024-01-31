namespace ArrayToExcel;

public class Text(string? value, bool wrap = false)
{
    public string? Value { get; } = value;
    public bool Wrap { get; } = wrap;
}
