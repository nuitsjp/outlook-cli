using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OutlookCli.OutlookCom;

internal sealed class OutlookInbox : IDisposable
{
    private const int InboxFolder = 6; // OlDefaultFolders.olFolderInbox
    internal const int MailItemClass = 43; // OlObjectClass.olMail
    private const int ShutdownTimeoutMilliseconds = 20_000;

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
        var applicationType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("クラシック Outlook の COM 登録が見つかりません。");

        using var currentProcess = Process.GetCurrentProcess();
        var sessionId = currentProcess.SessionId;
        using var existingOutlook = FindOutlookProcess(sessionId);
        application = Activator.CreateInstance(applicationType)!;
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
            Release(item);
            throw;
        }
    }

    public object? GetReceivedItem(string entryId) => ((dynamic)session!).GetItemFromID(entryId);

    public void Dispose()
    {
        using var outlookToWaitFor = startedOutlook;
        Release(folder);
        Release(session);
        Release(application);

        // Outlook started by COM can exit after the last reference is released.
        // Quit would also close a window the user opened during the query.
        if (outlookToWaitFor is not null && !outlookToWaitFor.WaitForExit(ShutdownTimeoutMilliseconds))
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

    internal static void Release(object? instance)
    {
        if (instance is not null)
        {
            Marshal.ReleaseComObject(instance);
        }
    }
}
