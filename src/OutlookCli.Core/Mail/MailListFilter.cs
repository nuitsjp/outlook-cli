namespace OutlookCli.Core.Mail;

public sealed record MailListFilter(
    bool UnreadOnly,
    string? FolderId,
    string? Subject,
    DateTimeOffset? Since);
