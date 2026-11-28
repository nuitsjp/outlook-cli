namespace OutlookCli.OutlookCom;

internal sealed class ComCleanup
{
    private readonly List<Exception> failures = [];
    private readonly Action<object> release;

    internal ComCleanup()
        : this(OutlookComObject.Release)
    {
    }

    internal ComCleanup(Action<object> release)
    {
        if (release is null)
        {
            throw new ArgumentNullException(nameof(release));
        }

        this.release = release;
    }

    internal IReadOnlyList<Exception> Failures => failures;

    internal bool MailSubmitted { get; private set; }

    internal static T Run<T>(Func<ComCleanup, T> work)
    {
        var cleanup = new ComCleanup();
        Exception? primary = null;
        try
        {
            return work(cleanup);
        }
        catch (Exception exception)
        {
            primary = exception;
            throw;
        }
        finally
        {
            cleanup.ThrowIfCleanupFailed(primary);
        }
    }

    internal static void Run(Action<ComCleanup> work)
    {
        Run(cleanup =>
        {
            work(cleanup);
            return true;
        });
    }

    internal static IEnumerable<T> Enumerate<T>(Func<ComCleanup, IEnumerable<T>> source)
    {
        var cleanup = new ComCleanup();
        Exception? primary = null;
        IEnumerator<T>? enumerator = null;
        try
        {
            while (true)
            {
                T current;
                try
                {
                    enumerator ??= source(cleanup).GetEnumerator();
                    if (!enumerator.MoveNext() || cleanup.failures.Count > 0)
                    {
                        break;
                    }
                    current = enumerator.Current;
                }
                catch (Exception exception)
                {
                    primary = exception;
                    throw;
                }

                yield return current;
            }
        }
        finally
        {
            if (enumerator is not null)
            {
                cleanup.Attempt(enumerator.Dispose);
            }

            cleanup.ThrowIfCleanupFailed(primary);
        }
    }

    internal void Release(object? value)
    {
        if (value is null)
        {
            return;
        }

        try
        {
            release(value);
        }
        catch (Exception exception)
        {
            failures.Add(exception);
        }
    }

    internal void Attempt(Action cleanup)
    {
        try
        {
            cleanup();
        }
        catch (Exception exception)
        {
            failures.Add(exception);
        }
    }

    internal void MarkMailSubmitted() => MailSubmitted = true;

    private void ThrowIfCleanupFailed(Exception? primary)
    {
        if (failures.Count > 0)
        {
            throw primary is null
                ? new OutlookComCleanupException(null, failures, MailSubmitted)
                : OutlookComCleanupException.Combine(primary, failures, MailSubmitted);
        }
    }
}
