namespace OutlookCli.Core.Mail;

public sealed record MailReplyResult(
    string EntryId,
    string DraftEntryId,
    string Subject);
