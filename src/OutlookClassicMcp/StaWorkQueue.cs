using System.Collections.Concurrent;

namespace OutlookClassicMcp;

/// <summary>
/// The one thread that is allowed to touch Outlook. Outlook's COM objects must be used from
/// a single STA thread, so every request is queued here and executed in order.
/// </summary>
public sealed class StaWorkQueue : IWorkQueue, IDisposable
{
    private readonly BlockingCollection<Action> _queue = [];

    public StaWorkQueue()
    {
        var thread = new Thread(Loop) { IsBackground = true, Name = "outlook-com" };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    public Task<T> RunAsync<T>(Func<T> work)
    {
        var done = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(() =>
        {
            try
            {
                done.SetResult(work());
            }
            catch (Exception ex)
            {
                done.SetException(ex);
            }
        });
        return done.Task;
    }

    private void Loop()
    {
        foreach (var action in _queue.GetConsumingEnumerable())
        {
            action();
        }
    }

    public void Dispose() => _queue.CompleteAdding();
}
