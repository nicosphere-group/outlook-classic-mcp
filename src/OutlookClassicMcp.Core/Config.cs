namespace OutlookClassicMcp;

/// <summary>Settings. Everything has a safe default, so no configuration is required.</summary>
public sealed record Config
{
    public const string AppDirName = "outlook-classic-mcp";
    public const int DefaultConfirmTimeout = 180;
    public const int MinConfirmTimeout = 15;
    public const int MaxConfirmTimeout = 600;

    public bool AllowSend { get; init; } = true;
    public bool Diagnostics { get; init; }
    public int ConfirmTimeout { get; init; } = DefaultConfirmTimeout;
    public string SaveDir { get; init; } = Path.Combine(Home, "Downloads", AppDirName);
    public IReadOnlyList<string> AttachRoots { get; init; } = DefaultAttachRoots();
    public string AuditDir { get; init; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppDirName);
    public string Lang { get; init; } = "auto"; // "ja" | "en" | "auto"

    private static string Home => Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

    private static string[] DefaultAttachRoots() => new[]
    {
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
        Path.Combine(Home, "Downloads"),
    }.Where(p => p.Length > 0).ToArray();

    public static Config FromEnvironment() => From(Environment.GetEnvironmentVariable);

    public static Config From(Func<string, string?> env)
    {
        // The extension host substitutes settings into the environment; a setting that was
        // left empty can arrive as an empty string or as an unreplaced "${...}" placeholder.
        string? Get(string name)
        {
            var value = env(name)?.Trim();
            return string.IsNullOrEmpty(value) || value.StartsWith("${", StringComparison.Ordinal)
                ? null
                : value;
        }

        var config = new Config();

        // Sending is on unless it is switched off explicitly.
        if (Get("OUTLOOK_MCP_ALLOW_SEND") is { } allow && IsOff(allow))
        {
            config = config with { AllowSend = false };
        }
        if (Get("OUTLOOK_MCP_DIAGNOSTICS") is { } diagnostics && IsOn(diagnostics))
        {
            config = config with { Diagnostics = true };
        }
        if (Get("OUTLOOK_MCP_CONFIRM_TIMEOUT") is { } timeout && int.TryParse(timeout, out var seconds))
        {
            config = config with
            {
                ConfirmTimeout = Math.Clamp(seconds, MinConfirmTimeout, MaxConfirmTimeout),
            };
        }
        if (Get("OUTLOOK_MCP_SAVE_DIR") is { } saveDir)
        {
            config = config with { SaveDir = saveDir };
        }
        if (Get("OUTLOOK_MCP_ATTACH_DIRS") is { } attachDirs)
        {
            var roots = attachDirs.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (roots.Length > 0)
            {
                config = config with { AttachRoots = roots };
            }
        }
        if (Get("OUTLOOK_MCP_AUDIT_DIR") is { } auditDir)
        {
            config = config with { AuditDir = auditDir };
        }
        if (Get("OUTLOOK_MCP_LANG")?.ToLowerInvariant() is "ja" or "en")
        {
            config = config with { Lang = Get("OUTLOOK_MCP_LANG")!.ToLowerInvariant() };
        }
        return config;
    }

    private static bool IsOff(string value) =>
        value is "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase);

    private static bool IsOn(string value) =>
        value is "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase);
}
