namespace OneDriveCopier;

/// <summary>
/// Lightweight CLI argument parser.
/// Supports:  --key value   and boolean  --flag
/// </summary>
internal sealed class ArgumentParser
{
    private readonly Dictionary<string, string> _values;
    private readonly HashSet<string> _flags;

    public ArgumentParser(string[] args)
    {
        _values = new(StringComparer.OrdinalIgnoreCase);
        _flags  = new(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            if (!arg.StartsWith("--")) continue;

            bool hasValue = i + 1 < args.Length && !args[i + 1].StartsWith("--");
            if (hasValue) _values[arg] = args[++i];
            else          _flags.Add(arg);
        }
    }

    public string? GetValue(string key) =>
        _values.TryGetValue(key, out var v) ? v : null;

    public bool HasFlag(string key) => _flags.Contains(key);
}
