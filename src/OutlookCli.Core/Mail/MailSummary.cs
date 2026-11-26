namespace OutlookCli.Core.Mail;

public sealed record MailSummary(string Subject, string SenderName, DateTimeOffset ReceivedAt, string EntryId);
