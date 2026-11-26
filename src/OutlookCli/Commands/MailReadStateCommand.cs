using System.CommandLine;
using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailReadStateCommand(IMailReadStateUpdater updater, TextWriter output, TextWriter error)
{
    public Command Create(bool unread)
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "一覧で取得したメールの Entry ID です。",
            Required = true,
        };
        var command = unread
            ? new Command("mark-unread", "指定したメールを未読にします。")
            : new Command("mark-read", "指定したメールを既読にします。");
        command.Options.Add(entryId);
        command.SetAction(result => Execute(result.GetValue(entryId), unread));
        return command;
    }

    private int Execute(string? entryId, bool unread)
    {
        if (entryId is null || string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にメールの Entry ID を指定してください。");
            return 1;
        }

        try
        {
            var change = updater.SetUnread(entryId, unread);
            output.WriteLine(JsonSerializer.Serialize(new
            {
                schemaVersion = 1,
                change.EntryId,
                unread,
                change.Changed,
            }, MailJson.Options));
            return 0;
        }
        catch (COMException exception)
        {
            var state = unread ? "未読" : "既読";
            error.WriteLine($"メールの{state}化に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
