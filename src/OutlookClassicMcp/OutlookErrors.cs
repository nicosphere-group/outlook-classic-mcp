using System.Runtime.InteropServices;

namespace OutlookClassicMcp;

/// <summary>Turns COM failures into messages a non-technical person can act on.</summary>
public static class OutlookErrors
{
    private const int ClassNotRegistered = unchecked((int)0x800401F3);
    private const int ClassNotRegistered2 = unchecked((int)0x80040154);
    private const int ServerExecutionFailed = unchecked((int)0x80080005);
    private const int RpcServerUnavailable = unchecked((int)0x800706BA);
    private const int RpcCallFailed = unchecked((int)0x800706BE);

    public static OutlookMcpException NotInstalled() => new OutlookUnavailableException(
        "Classic Outlook was not found on this PC. The new Outlook app is not supported; " +
        "classic Outlook (OUTLOOK.EXE) is required.");

    public static OutlookMcpException Translate(COMException ex) => ex.HResult switch
    {
        ClassNotRegistered or ClassNotRegistered2 => NotInstalled(),
        ServerExecutionFailed => new OutlookUnavailableException(
            "Outlook did not respond. Start Outlook normally (not 'Run as administrator'), " +
            "make sure Claude is not running as administrator either, and try again."),
        RpcServerUnavailable or RpcCallFailed => new OutlookUnavailableException(
            "The connection to Outlook was lost. Make sure Outlook is open and try again."),
        _ => new OutlookMcpException(
            string.IsNullOrWhiteSpace(ex.Message)
                ? "Outlook reported a problem and the operation was not completed."
                : $"Outlook reported a problem: {ex.Message}"),
    };
}
