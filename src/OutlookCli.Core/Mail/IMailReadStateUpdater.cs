namespace OutlookCli.Core.Mail;

public interface IMailReadStateUpdater
{
    MailReadStateChange SetUnread(string entryId, bool unread);
}
