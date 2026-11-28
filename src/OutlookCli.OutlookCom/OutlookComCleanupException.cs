namespace OutlookCli.OutlookCom;

public sealed class OutlookComCleanupException : Exception
{
    public OutlookComCleanupException(
        Exception? primaryException,
        IReadOnlyList<Exception> cleanupExceptions,
        bool mailSubmitted = false)
        : base(
            primaryException is null
                ? "Outlook COM リソースの解放に失敗しました。"
                : "Outlook の操作に失敗し、COM リソースの解放にも失敗しました。",
            primaryException)
    {
        PrimaryException = primaryException;
        CleanupExceptions = cleanupExceptions.ToArray();
        MailSubmitted = mailSubmitted;
        HResult = primaryException?.HResult ?? unchecked((int)0x80131500);
    }

    public Exception? PrimaryException { get; }

    public IReadOnlyList<Exception> CleanupExceptions { get; }

    public bool MailSubmitted { get; }

    public static OutlookComCleanupException Combine(
        Exception primaryException,
        IEnumerable<Exception> cleanupExceptions,
        bool mailSubmitted = false)
    {
        if (primaryException is OutlookComCleanupException existing)
        {
            return new OutlookComCleanupException(
                existing.PrimaryException,
                existing.CleanupExceptions.Concat(cleanupExceptions).ToArray(),
                existing.MailSubmitted || mailSubmitted);
        }

        return new OutlookComCleanupException(
            primaryException,
            cleanupExceptions.ToArray(),
            mailSubmitted);
    }
}
