using ModelContextProtocol;

namespace OutlookClassicMcp;

// McpException messages are passed to the client as-is, so every message here is
// written for a non-technical reader.

public class OutlookMcpException(string message) : McpException(message);

/// <summary>Outlook could not be reached (not installed, not responding, elevation mismatch).</summary>
public sealed class OutlookUnavailableException(string message) : OutlookMcpException(message);

/// <summary>A mail, folder, account or attachment does not exist (any more).</summary>
public sealed class NotFoundException(string message) : OutlookMcpException(message);

/// <summary>The item exists but is not an unsent draft.</summary>
public sealed class NotADraftException(string message) : OutlookMcpException(message);

/// <summary>Outside what this tool permits (path limits, disabled send).</summary>
public sealed class NotAllowedException(string message) : OutlookMcpException(message);

/// <summary>An argument could not be understood.</summary>
public sealed class InvalidInputException(string message) : OutlookMcpException(message);
