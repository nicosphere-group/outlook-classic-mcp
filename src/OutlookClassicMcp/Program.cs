using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookClassicMcp;

// Must be set before any window exists so the confirmation window is crisp on scaled displays.
System.Windows.Forms.Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
System.Windows.Forms.Application.EnableVisualStyles();
System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

var config = Config.FromEnvironment();
var builder = Host.CreateApplicationBuilder(args);

// stdout carries the MCP conversation; every log line must go to stderr.
builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

builder.Services.AddSingleton<IWorkQueue, StaWorkQueue>();
builder.Services.AddSingleton<IMailBackend, OutlookBackend>();
builder.Services.AddSingleton<IConfirmer>(new DialogConfirmer(config.Lang));
builder.Services.AddOutlookMcp(config);

await builder.Build().RunAsync();
