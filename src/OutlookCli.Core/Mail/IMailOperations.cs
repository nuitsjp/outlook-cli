namespace OutlookCli.Core.Mail;

public interface IMailOperations
{
    IEnumerable<MailFolderSummary> ListFolders();

    IEnumerable<MailAttachmentSummary> ListAttachments(string entryId);

    MailAttachmentSaveResult SaveAttachment(string entryId, string attachmentId, string path);

    MailMoveResult Move(string entryId, string folderId);

    MailReplyResult Reply(string entryId, string body);
}
