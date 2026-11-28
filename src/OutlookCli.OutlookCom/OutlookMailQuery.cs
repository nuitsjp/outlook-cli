using System.Globalization;
using OutlookCli.Core.Mail;

namespace OutlookCli.OutlookCom;

public sealed class OutlookMailQuery : IMailQuery, IMailOperations
{
    private const int MailItemType = 0; // OlItemType.olMailItem

    public IEnumerable<MailSummary> List(MailListFilter filter) =>
        ComCleanup.Enumerate(cleanup => ListCore(cleanup, filter));

    private static IEnumerable<MailSummary> ListCore(ComCleanup cleanup, MailListFilter filter)
    {
        using var connection = OutlookFolderConnection.Open(cleanup);
        object? selectedFolder = null;
        object? items = null;
        object? restrictedItems = null;
        var ownsSelectedFolder = false;

        try
        {
            selectedFolder = ResolveFolder(connection, filter.FolderId, out ownsSelectedFolder);
            items = ((dynamic)selectedFolder).Items;
            var restriction = filter.UnreadOnly ? "[UnRead] = True" : null;
            if (filter.Since is DateTimeOffset since)
            {
                // Outlook's DASL date literals are UTC and only support minute precision.
                var minute = since.UtcDateTime.ToString("g", CultureInfo.CurrentCulture).Replace("'", "''");
                restriction = $"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{minute}'";
                if (filter.UnreadOnly)
                {
                    restriction += " AND \"urn:schemas:httpmail:read\" = 0";
                }
            }

            if (restriction is not null)
            {
                restrictedItems = ((dynamic)items!).Restrict(restriction);
            }

            var results = (dynamic)(restrictedItems ?? items!);
            results.Sort("[ReceivedTime]", true);

            for (object? item = results.GetFirst(); item is not null; item = results.GetNext())
            {
                MailSummary summary;
                try
                {
                    if ((int)((dynamic)item).Class != OutlookFolderConnection.MailItemClass)
                    {
                        continue;
                    }

                    summary = ReadSummary(item);
                }
                finally
                {
                    cleanup.Release(item);
                }

                if (filter.Subject is not null
                    && summary.Subject.IndexOf(filter.Subject, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (filter.Since is DateTimeOffset receivedFrom && summary.ReceivedAt < receivedFrom)
                {
                    continue;
                }

                yield return summary;
            }
        }
        finally
        {
            cleanup.Release(restrictedItems);
            cleanup.Release(items);
            if (ownsSelectedFolder)
            {
                cleanup.Release(selectedFolder);
            }
        }
    }

    public MailDetail Get(string entryId) => ComCleanup.Run(cleanup =>
    {
        using var inbox = OutlookFolderConnection.Open(cleanup);
        var item = inbox.GetMail(entryId);
        try
        {
            var mail = (dynamic)item;
            return new MailDetail(
                ReadSummary(item),
                (string?)mail.To ?? string.Empty,
                (string?)mail.CC ?? string.Empty,
                (string?)mail.BCC ?? string.Empty,
                (string?)mail.Body ?? string.Empty);
        }
        finally
        {
            cleanup.Release(item);
        }
    });

    public IEnumerable<MailFolderSummary> ListFolders() =>
        ComCleanup.Enumerate(ListFoldersCore);

    private static IEnumerable<MailFolderSummary> ListFoldersCore(ComCleanup cleanup)
    {
        using var connection = OutlookFolderConnection.Open(cleanup);
        object? store = null;
        object? root = null;
        try
        {
            store = ((dynamic)connection.Folder).Store;
            root = ((dynamic)store!).GetRootFolder();
            foreach (var folder in ReadFolders(cleanup, root!, null, string.Empty))
            {
                yield return folder;
            }
        }
        finally
        {
            cleanup.Release(root);
            cleanup.Release(store);
        }
    }

    public IEnumerable<MailAttachmentSummary> ListAttachments(string entryId) =>
        ComCleanup.Enumerate(cleanup => ListAttachmentsCore(cleanup, entryId));

    private static IEnumerable<MailAttachmentSummary> ListAttachmentsCore(ComCleanup cleanup, string entryId)
    {
        using var connection = OutlookFolderConnection.Open(cleanup);
        var item = connection.GetMail(entryId);
        object? attachments = null;
        try
        {
            attachments = ((dynamic)item).Attachments;
            var collection = (dynamic)attachments!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? attachment = null;
                try
                {
                    attachment = collection.Item(index);
                    var value = (dynamic)attachment!;
                    yield return new MailAttachmentSummary(
                        index.ToString(CultureInfo.InvariantCulture),
                        (string?)value.FileName ?? string.Empty,
                        Convert.ToInt64(value.Size, CultureInfo.InvariantCulture),
                        Convert.ToInt32(value.Type, CultureInfo.InvariantCulture));
                }
                finally
                {
                    cleanup.Release(attachment);
                }
            }
        }
        finally
        {
            cleanup.Release(attachments);
            cleanup.Release(item);
        }
    }

    public MailAttachmentSaveResult SaveAttachment(string entryId, string attachmentId, string path) => ComCleanup.Run(cleanup =>
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("添付ファイルの保存先パスを指定してください。");
        }

        if (!int.TryParse(attachmentId, NumberStyles.None, CultureInfo.InvariantCulture, out var index)
            || index < 1)
        {
            throw new InvalidOperationException("添付ファイル ID には添付ファイル一覧の番号を指定してください。");
        }

        var fullPath = Path.GetFullPath(path);
        if (File.Exists(fullPath))
        {
            throw new InvalidOperationException($"保存先ファイルが既に存在します。上書きしません: {fullPath}");
        }

        using var connection = OutlookFolderConnection.Open(cleanup);
        var item = connection.GetMail(entryId);
        object? attachments = null;
        object? attachment = null;
        try
        {
            attachments = ((dynamic)item).Attachments;
            attachment = ((dynamic)attachments!).Item(index);
            AttachmentFileSaver.Save(fullPath, temporaryPath =>
                ((dynamic)attachment!).SaveAsFile(temporaryPath));

            return new MailAttachmentSaveResult(
                (string)((dynamic)item).EntryID,
                attachmentId,
                fullPath);
        }
        finally
        {
            cleanup.Release(attachment);
            cleanup.Release(attachments);
            cleanup.Release(item);
        }
    });

    public MailMoveResult Move(string entryId, string folderId) => ComCleanup.Run(cleanup =>
    {
        if (string.IsNullOrWhiteSpace(folderId))
        {
            throw new InvalidOperationException("移動先フォルダー ID を指定してください。");
        }

        using var connection = OutlookFolderConnection.Open(cleanup);
        var item = connection.GetMail(entryId);
        object? destination = null;
        object? parent = null;
        object? moved = null;
        try
        {
            destination = ResolveFolder(connection, folderId, out _);
            var destinationId = (string)((dynamic)destination!).EntryID;
            var storeId = (string)((dynamic)destination!).StoreID;
            parent = ((dynamic)item).Parent;
            var currentFolderId = (string?)((dynamic)parent!).EntryID ?? string.Empty;
            if (string.Equals(currentFolderId, destinationId, StringComparison.OrdinalIgnoreCase))
            {
                return new MailMoveResult((string)((dynamic)item).EntryID, destinationId, storeId, false);
            }

            moved = ((dynamic)item).Move(destination);
            var movedEntryId = (string)((dynamic)moved!).EntryID;
            return new MailMoveResult(movedEntryId, destinationId, storeId, true);
        }
        finally
        {
            cleanup.Release(moved);
            cleanup.Release(parent);
            cleanup.Release(destination);
            cleanup.Release(item);
        }
    });

    public MailReplyResult Reply(string entryId, string body) => ComCleanup.Run(cleanup =>
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            throw new InvalidOperationException("返信本文を指定してください。");
        }

        using var connection = OutlookFolderConnection.Open(cleanup);
        var item = connection.GetMail(entryId);
        object? reply = null;
        try
        {
            reply = ((dynamic)item).Reply();
            var draft = (dynamic)reply!;
            var quotedBody = (string?)draft.Body ?? string.Empty;
            draft.BodyFormat = 1; // OlBodyFormat.olFormatPlain
            draft.Body = string.IsNullOrEmpty(quotedBody)
                ? body
                : body + Environment.NewLine + Environment.NewLine + quotedBody;
            draft.Save();
            return new MailReplyResult(
                (string)((dynamic)item).EntryID,
                (string)draft.EntryID,
                (string?)draft.Subject ?? string.Empty);
        }
        finally
        {
            cleanup.Release(reply);
            cleanup.Release(item);
        }
    });

    internal static MailSummary ReadSummary(object item)
    {
        var mail = (dynamic)item;
        return new MailSummary(
            (string?)mail.Subject ?? string.Empty,
            (string?)mail.SenderName ?? string.Empty,
            new DateTimeOffset((DateTime)mail.ReceivedTime),
            (string)mail.EntryID);
    }

    private static object ResolveFolder(
        OutlookFolderConnection connection,
        string? folderId,
        out bool ownsFolder)
    {
        if (string.IsNullOrWhiteSpace(folderId))
        {
            ownsFolder = false;
            EnsureMailFolder(connection.Folder);
            return connection.Folder;
        }

        var storeId = (string)((dynamic)connection.Folder).StoreID;
        var folder = ((dynamic)connection.Session).GetFolderFromID(folderId, storeId);
        if (folder is null)
        {
            throw new InvalidOperationException("指定されたフォルダー ID に対応するフォルダーが見つかりません。");
        }

        ownsFolder = true;
        try
        {
            EnsureMailFolder(folder);
            return folder;
        }
        catch
        {
            ownsFolder = false;
            connection.Cleanup.Release(folder);
            throw;
        }
    }

    private static void EnsureMailFolder(object folder)
    {
        if ((int)((dynamic)folder).DefaultItemType != MailItemType)
        {
            throw new InvalidOperationException("指定されたフォルダーはメール フォルダーではありません。");
        }
    }

    private static IEnumerable<MailFolderSummary> ReadFolders(
        ComCleanup cleanup,
        object folder,
        string? parentFolderId,
        string parentPath)
    {
        var current = (dynamic)folder;
        var folderId = (string)current.EntryID;
        var displayName = (string?)current.Name ?? string.Empty;
        var storeId = (string?)current.StoreID ?? string.Empty;
        var path = string.IsNullOrEmpty(parentPath)
            ? displayName
            : parentPath + "\\" + displayName;
        if ((int)current.DefaultItemType == MailItemType)
        {
            yield return new MailFolderSummary(
                folderId,
                displayName,
                parentFolderId ?? string.Empty,
                path,
                storeId);
        }

        object? children = null;
        try
        {
            children = current.Folders;
            var collection = (dynamic)children!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? child = null;
                try
                {
                    child = collection.Item(index);
                    foreach (var result in ReadFolders(cleanup, child!, folderId, path))
                    {
                        yield return result;
                    }
                }
                finally
                {
                    cleanup.Release(child);
                }
            }
        }
        finally
        {
            cleanup.Release(children);
        }
    }

}
