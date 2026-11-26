namespace OutlookCli.Core.Mail;

public interface IMailQuery
{
    IEnumerable<MailSummary> List(bool unreadOnly);
    MailDetail Get(string entryId);
}
