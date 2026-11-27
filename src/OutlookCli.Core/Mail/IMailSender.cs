namespace OutlookCli.Core.Mail;

public interface IMailSender
{
    void Send(string to, string subject, string body);
}
