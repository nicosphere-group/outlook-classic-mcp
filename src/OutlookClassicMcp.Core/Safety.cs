using System.Text.RegularExpressions;

namespace OutlookClassicMcp;

/// <summary>File-path limits and pre-send warnings. No Outlook access here.</summary>
public static partial class Safety
{
    public const int ManyRecipients = 5;
    private const int MaxNameLength = 150;

    private static readonly HashSet<string> Reserved = new(
        new[] { "CON", "PRN", "AUX", "NUL" }
            .Concat(Enumerable.Range(1, 9).Select(i => $"COM{i}"))
            .Concat(Enumerable.Range(1, 9).Select(i => $"LPT{i}")),
        StringComparer.OrdinalIgnoreCase);

    [GeneratedRegex("[<>:\"/\\\\|?*\\x00-\\x1f]")]
    private static partial Regex BadChars();

    /// <summary>Turns an attacker-chosen attachment name into a harmless file name.</summary>
    public static string SanitizeFilename(string? name)
    {
        var clean = (name ?? "").Replace('\\', '/').Split('/')[^1];
        clean = BadChars().Replace(clean, "_").Trim(' ', '.');
        if (clean.Length == 0)
        {
            clean = "attachment";
        }
        if (Reserved.Contains(clean.Split('.')[0]))
        {
            clean = "_" + clean;
        }
        if (clean.Length > MaxNameLength)
        {
            var dot = clean.LastIndexOf('.');
            var ext = dot >= 0 ? clean[(dot + 1)..] : "";
            clean = dot >= 0 && ext.Length <= 10
                ? clean[..(MaxNameLength - ext.Length - 1)] + "." + ext
                : clean[..MaxNameLength];
        }
        return clean;
    }

    public static bool IsWithin(string path, IEnumerable<string> roots)
    {
        var full = Path.GetFullPath(path);
        foreach (var root in roots)
        {
            var rootFull = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));
            if (full.Equals(rootFull, StringComparison.OrdinalIgnoreCase) ||
                full.StartsWith(rootFull + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>A path inside <paramref name="saveDir"/> that does not exist yet. Never overwrites.</summary>
    public static string SafeSavePath(string saveDir, string? filename)
    {
        Directory.CreateDirectory(saveDir);
        var clean = SanitizeFilename(filename);
        var dot = clean.LastIndexOf('.');
        var stem = dot > 0 ? clean[..dot] : clean;
        var ext = dot > 0 ? clean[dot..] : "";
        var candidate = Path.Combine(saveDir, clean);
        for (var n = 1; File.Exists(candidate) || Directory.Exists(candidate); n++)
        {
            candidate = Path.Combine(saveDir, $"{stem} ({n}){ext}");
        }
        if (!IsWithin(candidate, [saveDir]))
        {
            throw new NotAllowedException("The attachment could not be saved inside the allowed folder.");
        }
        return Path.GetFullPath(candidate);
    }

    /// <summary>Files may be attached only from the allowed folders (limits data exfiltration).</summary>
    public static IReadOnlyList<string> CheckAttachSources(
        IEnumerable<string>? paths, IReadOnlyList<string> roots)
    {
        var result = new List<string>();
        foreach (var raw in paths ?? [])
        {
            if (!Path.IsPathFullyQualified(raw))
            {
                throw new NotAllowedException($"Give the full path of the file to attach: {raw}");
            }
            var full = Path.GetFullPath(raw);
            // A shortcut-like link inside an allowed folder must not reach outside it.
            if (File.Exists(full) && new FileInfo(full).ResolveLinkTarget(returnFinalTarget: true) is { } target)
            {
                full = target.FullName;
            }
            if (!IsWithin(full, roots))
            {
                throw new NotAllowedException(
                    $"Files can be attached only from these folders: {string.Join(", ", roots)}. " +
                    $"Not allowed: {raw}");
            }
            if (!File.Exists(full))
            {
                throw new NotFoundException($"The file to attach was not found: {raw}");
            }
            result.Add(full);
        }
        return result;
    }

    private static string Domain(string address)
    {
        var at = address.LastIndexOf('@');
        return at < 0 ? "" : address[(at + 1)..].Trim().ToLowerInvariant();
    }

    /// <summary>Things the confirmation window highlights. Codes are localized by the window.</summary>
    public static IReadOnlyList<Warning> ComputeWarnings(DraftSnapshot snapshot, bool replyAll = false)
    {
        var warnings = new List<Warning>();
        var recipients = snapshot.Recipients();
        if (recipients.Count >= ManyRecipients)
        {
            warnings.Add(new Warning("many_recipients", Count: recipients.Count));
        }
        if (replyAll)
        {
            warnings.Add(new Warning("reply_all"));
        }
        var own = Domain(snapshot.SenderAddress);
        if (own.Length > 0)
        {
            var external = recipients
                .Select(r => Domain(r.Address))
                .Where(d => d.Length > 0 && d != own)
                .Distinct()
                .Order(StringComparer.Ordinal)
                .ToList();
            if (external.Count > 0)
            {
                warnings.Add(new Warning("external", Domains: external));
            }
        }
        if (snapshot.Attachments.Count > 0)
        {
            warnings.Add(new Warning("attachments", Count: snapshot.Attachments.Count));
        }
        return warnings;
    }
}
