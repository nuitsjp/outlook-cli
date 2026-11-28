using System.Diagnostics;

namespace OutlookCli.OutlookCom;

internal sealed class OutlookApplication : IDisposable
{
    private const int ShutdownTimeoutMilliseconds = 20_000;

    private readonly ComCleanup cleanup;
    private readonly Process? startedOutlook;
    private bool disposed;

    internal OutlookApplication(ComCleanup cleanup, object instance, Process? startedOutlook)
    {
        this.cleanup = cleanup;
        Instance = instance;
        this.startedOutlook = startedOutlook;
    }

    internal object Instance { get; }

    internal static OutlookApplication Open(ComCleanup cleanup)
    {
        using var caller = Process.GetCurrentProcess();
        var sessionId = caller.SessionId;
        using var existingOutlook = FindOutlookProcess(sessionId);
        var instance = OutlookComObject.CreateApplication();
        try
        {
            var startedOutlook = existingOutlook is null ? FindOutlookProcess(sessionId) : null;
            return new OutlookApplication(cleanup, instance, startedOutlook);
        }
        catch
        {
            cleanup.Release(instance);
            throw;
        }
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        cleanup.Release(Instance);
        if (startedOutlook is null)
        {
            return;
        }

        // Quit would also close a window the user opened during the operation.
        cleanup.Attempt(() =>
        {
            if (!startedOutlook.WaitForExit(ShutdownTimeoutMilliseconds))
            {
                throw new InvalidOperationException(
                    $"COM 参照の解放後も Outlook (PID {startedOutlook.Id}) が終了しませんでした。");
            }
        });
        cleanup.Attempt(startedOutlook.Dispose);
    }

    private static Process? FindOutlookProcess(int sessionId)
    {
        var processes = Process.GetProcessesByName("OUTLOOK");
        Process? match = null;
        try
        {
            match = Array.Find(processes, process => process.SessionId == sessionId);
            return match;
        }
        finally
        {
            foreach (var process in processes)
            {
                if (process != match)
                {
                    process.Dispose();
                }
            }
        }
    }
}
