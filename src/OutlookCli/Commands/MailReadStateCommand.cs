using System.CommandLine;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailReadStateCommand(IMailReadStateUpdater updater, JsonOutput json, TextWriter error)
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

        return CommandExecution.Run(error, unread ? "メールの未読化" : "メールの既読化", () =>
        {
            var change = updater.SetUnread(entryId, unread);
            json.WriteLine(new
            {
                schemaVersion = 1,
                change.EntryId,
                unread,
                change.Changed,
            });
        });
    }
}
