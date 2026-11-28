namespace OutlookCli.Core.Mail;

public sealed record MailMoveResult(
    string EntryId,
    string FolderId,
    string StoreId,
    bool Changed);
