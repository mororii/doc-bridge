using System.Collections.Concurrent;

namespace DocBridge.Core.Services;

/// <summary>
/// COM 호출 전용 STA 스레드 디스패처.
/// Office/한글/AutoCAD COM은 STA에서 호출하는 것이 안전하다.
/// 모든 COM 접근은 이 러너를 통해 직렬화한다.
/// </summary>
public sealed class StaThreadRunner : IDisposable
{
    private sealed record WorkItem(Func<object?> Work, TaskCompletionSource<object?> Tcs);

    private readonly BlockingCollection<WorkItem> _queue = new();
    private readonly Thread _thread;
    private bool _disposed;

    public StaThreadRunner(string name)
    {
        _thread = new Thread(Run) { IsBackground = true, Name = name };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
    }

    private void Run()
    {
        foreach (var item in _queue.GetConsumingEnumerable())
        {
            try { item.Tcs.TrySetResult(item.Work()); }
            catch (Exception ex) { item.Tcs.TrySetException(ex); }
        }
    }

    public T Invoke<T>(Func<T> work, TimeSpan? timeout = null)
    {
        if (_disposed) throw new ObjectDisposedException(nameof(StaThreadRunner));
        var tcs = new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
        _queue.Add(new WorkItem(() => work(), tcs));
        if (timeout is null)
            return (T)tcs.Task.GetAwaiter().GetResult()!;
        if (tcs.Task.Wait(timeout.Value))
            return (T)tcs.Task.GetAwaiter().GetResult()!;
        throw new TimeoutException($"STA work item did not complete within {timeout.Value.TotalSeconds}s (possible COM modal dialog)");
    }

    public void Invoke(Action work) => Invoke<object?>(() => { work(); return null; });

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _queue.CompleteAdding();
        if (!_thread.Join(TimeSpan.FromSeconds(5)))
        {
            // STA 스레드가 COM 호출에서 멈춘 경우 프로세스 종료를 막지 않도록 background 스레드로 둠
        }
        _queue.Dispose();
    }
}
