namespace OutlookCli.OutlookCom;

internal sealed class OutlookFolderConnection(
    ComCleanup cleanup, OutlookApplication application, object session, object folder) : IDisposable
{
    private const int InboxFolder = 6; // OlDefaultFolders.olFolderInbox
    internal const int MailItemClass = 43; // OlObjectClass.olMail
    private bool disposed;

    internal ComCleanup Cleanup { get; } = cleanup;
    public object Application => application.Instance;
    public object Session { get; } = session;
    public object Folder { get; } = folder;

    internal static OutlookFolderConnection Open(
        ComCleanup cleanup, int defaultFolder = InboxFolder, string? owner = null) =>
        Open(cleanup, OutlookApplication.Open(cleanup), defaultFolder, owner);

    internal static OutlookFolderConnection Open(
        ComCleanup cleanup, OutlookApplication application, int defaultFolder, string? owner)
    {
        object? session = null;
        try
        {
            session = ((dynamic)application.Instance).GetNamespace("MAPI");
            object folder;
            if (owner is null)
            {
                folder = ((dynamic)session).GetDefaultFolder(defaultFolder);
            }
            else
            {
                var recipient = ResolveRecipient(cleanup, session, owner);
                try
                {
                    folder = ((dynamic)session).GetSharedDefaultFolder((dynamic)recipient, defaultFolder);
                }
                finally
                {
                    cleanup.Release(recipient);
                }
            }

            return new OutlookFolderConnection(cleanup, application, session, folder);
        }
        catch
        {
            cleanup.Release(session);
            application.Dispose();
            throw;
        }
    }

    public object ResolveRecipient(string name) => ResolveRecipient(Cleanup, Session, name);

    private static object ResolveRecipient(ComCleanup cleanup, object session, string name)
    {
        object recipient = ((dynamic)session).CreateRecipient(name);
        try
        {
            if (!(bool)((dynamic)recipient).Resolve())
            {
                throw new InvalidOperationException($"Outlook で宛先を解決できませんでした。{name}");
            }

            return recipient;
        }
        catch
        {
            cleanup.Release(recipient);
            throw;
        }
    }

    public object GetMail(string entryId)
    {
        var item = GetItem(entryId);
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
            Cleanup.Release(item);
            throw;
        }
    }

    public object? GetReceivedItem(string entryId) => ((dynamic)Session).GetItemFromID(entryId);

    public object? GetItem(string entryId)
    {
        var storeId = (string)((dynamic)Folder).StoreID;
        return ((dynamic)Session).GetItemFromID(entryId, storeId);
    }

    public void Dispose()
    {
        if (disposed)
        {
            return;
        }

        disposed = true;
        Cleanup.Release(Folder);
        Cleanup.Release(Session);
        application.Dispose();
    }
}
