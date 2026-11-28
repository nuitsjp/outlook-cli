namespace OutlookCli.Core.Mail;

public sealed record MailAttachmentSummary(
    string AttachmentId,
    string FileName,
    long Size,
    int Type);
