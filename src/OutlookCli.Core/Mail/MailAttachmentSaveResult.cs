namespace OutlookCli.Core.Mail;

public sealed record MailAttachmentSaveResult(
    string EntryId,
    string AttachmentId,
    string Path);
