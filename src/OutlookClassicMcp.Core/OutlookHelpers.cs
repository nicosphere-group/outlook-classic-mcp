using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookClassicMcp;

/// <summary>Outlook-related logic that needs no Outlook, so it can be unit-tested anywhere.</summary>
public static partial class OutlookHelpers
{
    private static readonly string[] SearchProperties =
    [
        "urn:schemas:httpmail:subject",
        "urn:schemas:httpmail:fromname",
        "urn:schemas:httpmail:fromemail",
        "urn:schemas:httpmail:textdescription",
    ];

    public static string StoreKey(string storeId) =>
        Convert.ToHexString(SHA1.HashData(Encoding.UTF8.GetBytes(storeId)))[..10].ToLowerInvariant();

    /// <summary>
    /// An email is identified by its EntryID together with its StoreID. The StoreID is long,
    /// so it is carried as a short key that is resolved against the stores at lookup time.
    /// </summary>
    public static string MakeId(string storeId, string entryId) => $"{StoreKey(storeId)}.{entryId}";

    public static (string StoreKey, string EntryId) SplitId(string? itemId)
    {
        var text = (itemId ?? "").Trim();
        var dot = text.IndexOf('.');
        if (dot <= 0 || dot == text.Length - 1)
        {
            throw new InvalidInputException("That email reference is not valid. List the emails again.");
        }
        return (text[..dot], text[(dot + 1)..]);
    }

    /// <summary>
    /// Outlook hands back local wall-clock time without a zone. Attach the local offset.
    /// Outlook's "no date" marker (year 4501) becomes null.
    /// </summary>
    public static DateTimeOffset? ToLocal(DateTime value)
    {
        if (value.Year > 4000 || value.Year < 1900)
        {
            return null;
        }
        var local = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        return new DateTimeOffset(local, TimeZoneInfo.Local.GetUtcOffset(local));
    }

    public static string Iso(DateTimeOffset? value) => value?.ToString("yyyy-MM-ddTHH:mm:sszzz") ?? "";

    public static string DaslLike(string property, string text) =>
        $"\"{property}\" like '%{text.Replace("'", "''")}%'";

    /// <summary>
    /// Text conditions only. Dates are deliberately not part of the filter: Outlook's date
    /// filter syntax depends on the PC's regional settings, so dates are compared in code.
    /// </summary>
    public static string? BuildSearchFilter(SearchTerms terms)
    {
        var clauses = new List<string>();
        if (terms.Query is { } query)
        {
            clauses.Add("(" + string.Join(" OR ", SearchProperties.Select(p => DaslLike(p, query))) + ")");
        }
        if (terms.Subject is { } subject)
        {
            clauses.Add(DaslLike(SearchProperties[0], subject));
        }
        if (terms.Sender is { } sender)
        {
            clauses.Add($"({DaslLike(SearchProperties[1], sender)} OR {DaslLike(SearchProperties[2], sender)})");
        }
        if (terms.Body is { } body)
        {
            clauses.Add(DaslLike(SearchProperties[3], body));
        }
        return clauses.Count == 0 ? null : "@SQL=" + string.Join(" AND ", clauses);
    }

    [GeneratedRegex("<body[^>]*>", RegexOptions.IgnoreCase)]
    private static partial Regex BodyTag();

    /// <summary>Puts plain text above the quoted original without destroying its formatting.</summary>
    public static string InsertHtmlReply(string? htmlBody, string text)
    {
        var html = htmlBody ?? "";
        var lines = text.ReplaceLineEndings("\n").Split('\n').Select(WebUtility.HtmlEncode);
        var block = "<div>" + string.Join("<br>", lines) + "</div><br>";
        var match = BodyTag().Match(html);
        return match.Success
            ? html[..(match.Index + match.Length)] + block + html[(match.Index + match.Length)..]
            : block + html;
    }
}
