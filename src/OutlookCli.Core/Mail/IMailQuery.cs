namespace OutlookCli.Core.Mail;

public interface IMailQuery
{
    IEnumerable<MailSummary> List(MailListFilter filter);
    MailDetail Get(string entryId);
}
