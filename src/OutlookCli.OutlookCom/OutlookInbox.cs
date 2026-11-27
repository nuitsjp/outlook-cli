using System.Diagnostics;

namespace OutlookCli.OutlookCom;

internal sealed class OutlookInbox : IDisposable
{
    private const int InboxFolder = 6; // OlDefaultFolders.olFolderInbox
    internal const int MailItemClass = 43; // OlObjectClass.olMail
    private static readonly TimeSpan ShutdownTimeout = TimeSpan.FromSeconds(20);

    private Process? startedOutlook;
    private object? application;
    private object? session;
    private object? folder;

    private OutlookInbox()
    {
    }

    public object Application => application!;
    public object Folder => folder!;

    public static OutlookInbox Open()
    {
        var inbox = new OutlookInbox();
        try
        {
            inbox.Connect();
            return inbox;
        }
        catch
        {
            inbox.Dispose();
            throw;
        }
    }

    private void Connect()
    {
        using var currentProcess = Process.GetCurrentProcess();
        var sessionId = currentProcess.SessionId;
        using var existingOutlook = FindOutlookProcess(sessionId);
        application = OutlookComObject.CreateApplication();
        if (existingOutlook is null)
        {
            startedOutlook = FindOutlookProcess(sessionId);
        }

        session = ((dynamic)application).GetNamespace("MAPI");
        folder = ((dynamic)session!).GetDefaultFolder(InboxFolder);
    }

    public object GetMail(string entryId)
    {
        var storeId = (string)((dynamic)Folder).StoreID;
        object? item = ((dynamic)session!).GetItemFromID(entryId, storeId);
        try
        {
            if (item is null || (int)((dynamic)item).Class != MailItemClass)
            {
                throw new InvalidOperationException("指定された ID に対応するメールが見つかりません。");
            }

            return item;
        }
        catch
        {
            OutlookComObject.Release(item);
            throw;
        }
    }

    public object? GetReceivedItem(string entryId) => ((dynamic)session!).GetItemFromID(entryId);

    public void Dispose()
    {
        using var outlookToWaitFor = startedOutlook;
        OutlookComObject.Release(folder);
        OutlookComObject.Release(session);
        OutlookComObject.Release(application);

        // Outlook started by COM can exit after the last reference is released.
        // Quit would also close a window the user opened during the query.
        if (outlookToWaitFor is not null && !outlookToWaitFor.WaitForExit(ShutdownTimeout))
        {
            throw new InvalidOperationException(
                $"COM 参照の解放後も Outlook (PID {outlookToWaitFor.Id}) が終了しませんでした。");
        }
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
