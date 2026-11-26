namespace OutlookCli.Core.Mail;

public sealed record MailDetail(MailSummary Summary, string To, string Cc, string Bcc, string Body);
