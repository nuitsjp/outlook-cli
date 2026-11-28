namespace OutlookCli.Core.Mail;

public sealed record MailFolderSummary(
    string FolderId,
    string DisplayName,
    string ParentFolderId,
    string Path,
    string StoreId);
