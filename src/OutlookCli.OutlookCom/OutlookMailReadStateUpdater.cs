using OutlookCli.Core.Mail;

namespace OutlookCli.OutlookCom;

public sealed class OutlookMailReadStateUpdater : IMailReadStateUpdater
{
    public MailReadStateChange SetUnread(string entryId, bool unread)
    {
        using var inbox = OutlookInbox.Open();
        var item = inbox.GetMail(entryId);
        try
        {
            var mail = (dynamic)item;
            var change = new MailReadStateChange((string)mail.EntryID, (bool)mail.UnRead != unread);
            if (change.Changed)
            {
                mail.UnRead = unread;
                mail.Save();
            }

            return change;
        }
        finally
        {
            OutlookComObject.Release(item);
        }
    }
}
