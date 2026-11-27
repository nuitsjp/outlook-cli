using OutlookCli.Core.Mail;

namespace OutlookCli.OutlookCom;

public sealed class OutlookMailQuery : IMailQuery
{
    public IEnumerable<MailSummary> List(bool unreadOnly)
    {
        using var inbox = OutlookInbox.Open();
        object? items = null;
        object? restrictedItems = null;

        try
        {
            items = ((dynamic)inbox.Folder).Items;
            if (unreadOnly)
            {
                restrictedItems = ((dynamic)items!).Restrict("[UnRead] = True");
            }

            var results = (dynamic)(restrictedItems ?? items!);
            results.Sort("[ReceivedTime]", true);

            for (object? item = results.GetFirst(); item is not null; item = results.GetNext())
            {
                MailSummary summary;
                try
                {
                    if ((int)((dynamic)item).Class != OutlookInbox.MailItemClass)
                    {
                        continue;
                    }

                    summary = ReadSummary(item);
                }
                finally
                {
                    OutlookComObject.Release(item);
                }

                yield return summary;
            }
        }
        finally
        {
            OutlookComObject.Release(restrictedItems);
            OutlookComObject.Release(items);
        }
    }

    public MailDetail Get(string entryId)
    {
        using var inbox = OutlookInbox.Open();
        var item = inbox.GetMail(entryId);
        try
        {
            var mail = (dynamic)item;
            return new MailDetail(
                ReadSummary(item),
                (string?)mail.To ?? string.Empty,
                (string?)mail.CC ?? string.Empty,
                (string?)mail.BCC ?? string.Empty,
                (string?)mail.Body ?? string.Empty);
        }
        finally
        {
            OutlookComObject.Release(item);
        }
    }

    internal static MailSummary ReadSummary(object item)
    {
        var mail = (dynamic)item;
        return new MailSummary(
            (string)mail.Subject,
            (string)mail.SenderName,
            new DateTimeOffset((DateTime)mail.ReceivedTime),
            (string)mail.EntryID);
    }
}
