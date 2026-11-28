using System.Runtime.InteropServices;
using OutlookCli.OutlookCom;

namespace OutlookCli.Commands;

internal static class CommandExecution
{
    public static int Run(TextWriter error, string operation, Action action)
    {
        try
        {
            action();
            return 0;
        }
        catch (OutlookComCleanupException exception)
        {
            if (exception.PrimaryException is { } primary)
            {
                ReportFailure(error, operation, primary);
            }

            ReportCleanupFailures(error, exception);
        }
        catch (Exception exception) when (exception is COMException or InvalidOperationException or IOException)
        {
            ReportFailure(error, operation, exception);
        }

        return 1;
    }

    internal static int RunList<T>(TextWriter error, string operation, Func<IEnumerable<T>> read, Action<T> writeRow)
    {
        return Run(error, operation, () =>
        {
            var items = read().GetEnumerator();
            Exception? primary = null;
            try
            {
                while (items.MoveNext())
                {
                    writeRow(items.Current);
                }
            }
            catch (Exception exception)
            {
                primary = exception;
                throw;
            }
            finally
            {
                try
                {
                    items.Dispose();
                }
                catch (OutlookComCleanupException cleanup) when (primary is not null)
                {
                    throw OutlookComCleanupException.Combine(primary, cleanup.CleanupExceptions, cleanup.MailSubmitted);
                }
            }
        });
    }

    internal static void ReportCleanupFailures(TextWriter error, OutlookComCleanupException exception)
    {
        foreach (var failure in exception.CleanupExceptions)
        {
            error.WriteLine($"Outlook の後処理に失敗しました。{Diagnostic(failure)}");
        }
    }

    private static void ReportFailure(TextWriter error, string operation, Exception exception)
    {
        error.WriteLine(exception is COMException
            ? $"{operation}に失敗しました {Diagnostic(exception)}"
            : exception.Message);
    }

    private static string Diagnostic(Exception exception) => exception is COMException
        ? $"(HRESULT 0x{exception.HResult:X8})。{exception.Message}"
        : exception.Message;
}
