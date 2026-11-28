using System.Runtime.InteropServices;
using System.CommandLine;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailFolderCommand(IMailOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("folder", "メール フォルダーを操作します。");
        command.Subcommands.Add(CreateList());
        return command;
    }

    private Command CreateList()
    {
        var command = new Command("list", "既定ストアのメール フォルダーを取得します。");
        command.SetAction(_ => ExecuteList());
        return command;
    }

    private int ExecuteList()
    {
        try
        {
            foreach (var folder in operations.ListFolders())
            {
                json.WriteLine(new
                {
                    schemaVersion = 1,
                    folder.FolderId,
                    folder.DisplayName,
                    folder.ParentFolderId,
                    folder.Path,
                    folder.StoreId,
                });
            }

            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"メール フォルダー一覧の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
