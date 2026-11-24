using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using OutlookClassicMcp;
using OutlookClassicMcp.Tests;

// SMOKE_DECISION chooses what the pretend human answers at the confirmation window.
var decision = Enum.Parse<Decision>(Environment.GetEnvironmentVariable("SMOKE_DECISION") ?? "Cancel", true);
var backend = new FakeBackend();
backend.AddInbox("store.mail1", subject: "【至急】お見積りの件①", body: "髙橋様　～よろしく～");

var builder = Host.CreateApplicationBuilder(args);
builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);
builder.Services.AddSingleton<IWorkQueue, ThreadPoolQueue>();
builder.Services.AddSingleton<IMailBackend>(backend);
builder.Services.AddSingleton<IConfirmer>(new FakeConfirmer(decision));
builder.Services.AddOutlookMcp(Config.FromEnvironment());
await builder.Build().RunAsync();

internal sealed class ThreadPoolQueue : IWorkQueue
{
    public Task<T> RunAsync<T>(Func<T> work) => Task.Run(work);
}
