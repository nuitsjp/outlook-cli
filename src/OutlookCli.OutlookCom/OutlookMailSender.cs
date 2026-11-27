using OutlookCli.Core.Mail;

namespace OutlookCli.OutlookCom;

public sealed class OutlookMailSender : IMailSender
{
    private const int MailItemType = 0; // OlItemType.olMailItem
    private const int PlainText = 1; // OlBodyFormat.olFormatPlain

    public void Send(string to, string subject, string body)
    {
        var application = OutlookComObject.CreateApplication();
        object? item = null;
        object? recipients = null;
        object? recipient = null;
        try
        {
            item = ((dynamic)application).CreateItem(MailItemType);
            var mail = (dynamic)item!;
            recipients = mail.Recipients;
            recipient = ((dynamic)recipients!).Add(to);
            if (!(bool)((dynamic)recipient!).Resolve())
            {
                throw new InvalidOperationException("Outlook で宛先を解決できませんでした。");
            }

            mail.Subject = subject;
            mail.BodyFormat = PlainText;
            mail.Body = body;
            mail.Send();
        }
        finally
        {
            OutlookComObject.Release(recipient);
            OutlookComObject.Release(recipients);
            OutlookComObject.Release(item);
            OutlookComObject.Release(application);
        }
    }
}
