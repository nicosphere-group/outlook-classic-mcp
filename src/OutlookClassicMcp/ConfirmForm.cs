using System.Runtime.InteropServices;

namespace OutlookClassicMcp;

/// <summary>
/// The confirmation window. Sending is never the default: closing the window, pressing
/// Enter or Esc, or waiting all mean "do not send".
/// </summary>
internal sealed class ConfirmForm : Form
{
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 1000 };
    private readonly Label _countdown = new() { AutoSize = true, ForeColor = Color.DimGray };
    private readonly string _countdownFormat;
    private int _remaining;

    public Decision Decision { get; private set; } = Decision.Cancel;

    public ConfirmForm(DialogView view, int timeoutSeconds)
    {
        _countdownFormat = view.CountdownFormat;
        _remaining = timeoutSeconds;

        var baseFont = SystemFonts.MessageBoxFont ?? Font;
        var font = new Font(baseFont.FontFamily, 10f);
        var bold = new Font(font, FontStyle.Bold);

        Text = view.Title;
        Font = font;
        AutoScaleMode = AutoScaleMode.Dpi;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(720, 560);
        MinimumSize = new Size(560, 420);
        MinimizeBox = false;
        MaximizeBox = true;
        ShowIcon = false;
        ShowInTaskbar = true;
        TopMost = true;
        Padding = new Padding(14);

        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        layout.Controls.Add(new Label
        {
            Text = view.Heading,
            Font = new Font(font.FontFamily, 13f, FontStyle.Bold),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 8),
        }, 0, 0);

        if (view.WarningLines.Count > 0)
        {
            layout.Controls.Add(new Label
            {
                Text = string.Join(Environment.NewLine, view.WarningLines.Select(line => "⚠ " + line)),
                Font = bold,
                BackColor = Color.FromArgb(0xFF, 0xF4, 0xCE),
                ForeColor = Color.FromArgb(0x8A, 0x1F, 0x11),
                AutoSize = true,
                Dock = DockStyle.Fill,
                Padding = new Padding(10, 8, 10, 8),
                Margin = new Padding(0, 0, 0, 8),
            }, 0, 1);
        }

        var text = new RichTextBox
        {
            ReadOnly = true,
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            DetectUrls = false, // links in an email must not be clickable from here
            ShortcutsEnabled = true,
            TabStop = false,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            WordWrap = true,
        };
        foreach (var (label, value) in view.Rows)
        {
            Append(text, label + Environment.NewLine, bold, 0);
            Append(text, value + Environment.NewLine, font, 24);
        }
        Append(text, Environment.NewLine + view.BodyLabel + Environment.NewLine, bold, 0);
        Append(text, view.Body + Environment.NewLine, font, 0);
        text.Select(0, 0);
        layout.Controls.Add(text, 0, 2);

        var send = MakeButton(view.SendText, Decision.Send);
        var open = MakeButton(view.OpenText, Decision.Open);
        var cancel = MakeButton(view.CancelText, Decision.Cancel);

        var buttons = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 4,
            RowCount = 1,
            Margin = new Padding(0, 12, 0, 0),
        };
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        buttons.Controls.Add(send, 0, 0);
        buttons.Controls.Add(open, 1, 0);
        buttons.Controls.Add(cancel, 3, 0);
        layout.Controls.Add(buttons, 0, 3);

        _countdown.Margin = new Padding(0, 8, 0, 0);
        layout.Controls.Add(_countdown, 0, 4);
        Controls.Add(layout);

        // No AcceptButton: Enter never means "send". Focus starts on "Don't send", so Enter or
        // Space there, Esc, and closing the window all decline.
        AcceptButton = null;
        CancelButton = cancel;
        ActiveControl = cancel;

        _timer.Tick += (_, _) => Tick();
        UpdateCountdown();
    }

    private Button MakeButton(string text, Decision decision)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Padding = new Padding(14, 6, 14, 6),
            Margin = new Padding(0, 0, 8, 0),
        };
        button.Click += (_, _) =>
        {
            Decision = decision;
            Close();
        };
        return button;
    }

    private static void Append(RichTextBox box, string text, Font font, int indent)
    {
        box.SelectionStart = box.TextLength;
        box.SelectionLength = 0;
        box.SelectionFont = font;
        box.SelectionIndent = indent;
        box.AppendText(text);
    }

    private void Tick()
    {
        _remaining--;
        if (_remaining <= 0)
        {
            Decision = Decision.Timeout;
            Close();
            return;
        }
        UpdateCountdown();
    }

    private void UpdateCountdown() => _countdown.Text = string.Format(_countdownFormat, _remaining);

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        _timer.Start();
        Activate();
        FlashWindow(Handle, true);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _timer.Stop();
        _timer.Dispose();
        base.OnFormClosed(e);
    }

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FlashWindow(IntPtr hWnd, [MarshalAs(UnmanagedType.Bool)] bool bInvert);
}

/// <summary>Shows the confirmation window on its own thread and waits for the human.</summary>
public sealed class DialogConfirmer(string lang) : IConfirmer
{
    public Decision Ask(DraftSnapshot snapshot, IReadOnlyList<Warning> warnings, int timeoutSeconds)
    {
        var view = DialogText.Build(snapshot, warnings, lang);
        var decision = Decision.Failed;
        var thread = new Thread(() =>
        {
            try
            {
                using var form = new ConfirmForm(view, timeoutSeconds);
                System.Windows.Forms.Application.Run(form);
                decision = form.Decision;
            }
            catch (Exception)
            {
                decision = Decision.Failed; // never send when the window could not be shown
            }
        })
        { IsBackground = true, Name = "confirm-dialog" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();
        return decision;
    }
}
