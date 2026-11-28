using System.CommandLine;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailMoveCommand(IMailOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "移動するメールの Entry ID です。",
            Required = true,
        };
        var folder = new Option<string>("--to-folder")
        {
            Description = "移動先フォルダーの Entry ID です。",
            Required = true,
        };
        var command = new Command("move", "メールを指定したフォルダーへ移動します。");
        command.Options.Add(entryId);
        command.Options.Add(folder);
        command.SetAction(result => Execute(result.GetValue(entryId), result.GetValue(folder)));
        return command;
    }

    private int Execute(string? entryId, string? folder)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にメールの Entry ID を指定してください。");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(folder))
        {
            error.WriteLine("--to-folder に移動先フォルダーの Entry ID を指定してください。");
            return 1;
        }

        try
        {
            var result = operations.Move(entryId!, folder!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                result.EntryId,
                result.FolderId,
                result.StoreId,
                result.Changed,
            });
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"メールの移動に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
