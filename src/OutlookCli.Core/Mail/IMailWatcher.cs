namespace OutlookCli.Core.Mail;

public interface IMailWatcher
{
    void Watch(Action<MailSummary> onMail, Action onReady, CancellationToken cancellationToken);
}
