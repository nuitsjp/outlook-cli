using System.CommandLine;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailAttachmentCommand(IMailOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("attachment", "メールの添付ファイルを操作します。");
        command.Subcommands.Add(CreateList());
        command.Subcommands.Add(CreateSave());
        return command;
    }

    private Command CreateList()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "対象メールの Entry ID です。",
            Required = true,
        };
        var command = new Command("list", "メールの添付ファイル一覧を取得します。");
        command.Options.Add(entryId);
        command.SetAction(result => ExecuteList(result.GetValue(entryId)));
        return command;
    }

    private Command CreateSave()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "対象メールの Entry ID です。",
            Required = true,
        };
        var attachmentId = new Option<string>("--attachment-id")
        {
            Description = "添付ファイル一覧の Attachment ID です。",
            Required = true,
        };
        var path = new Option<string>("--path")
        {
            Description = "保存先のファイルパスです。既存ファイルは上書きしません。",
            Required = true,
        };
        var command = new Command("save", "添付ファイルを指定したパスへ保存します。");
        command.Options.Add(entryId);
        command.Options.Add(attachmentId);
        command.Options.Add(path);
        command.SetAction(result => ExecuteSave(
            result.GetValue(entryId),
            result.GetValue(attachmentId),
            result.GetValue(path)));
        return command;
    }

    private int ExecuteList(string? entryId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にメールの Entry ID を指定してください。");
            return 1;
        }

        try
        {
            foreach (var attachment in operations.ListAttachments(entryId!))
            {
                json.WriteLine(new
                {
                    schemaVersion = 1,
                    attachment.AttachmentId,
                    attachment.FileName,
                    attachment.Size,
                    attachment.Type,
                });
            }

            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"添付ファイル一覧の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }

    private int ExecuteSave(string? entryId, string? attachmentId, string? path)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にメールの Entry ID を指定してください。");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(attachmentId))
        {
            error.WriteLine("--attachment-id に添付ファイル一覧の Attachment ID を指定してください。");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(path))
        {
            error.WriteLine("--path に保存先のファイルパスを指定してください。");
            return 1;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(path);
        }
        catch (ArgumentException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }

        if (File.Exists(fullPath))
        {
            error.WriteLine($"保存先ファイルが既に存在します。上書きしません: {fullPath}");
            return 1;
        }

        try
        {
            var result = operations.SaveAttachment(entryId!, attachmentId!, fullPath);
            json.WriteLine(new
            {
                schemaVersion = 1,
                result.EntryId,
                result.AttachmentId,
                result.Path,
                saved = true,
            });
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"添付ファイルの保存に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
