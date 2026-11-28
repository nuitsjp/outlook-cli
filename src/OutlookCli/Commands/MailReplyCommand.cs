using System.CommandLine;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailReplyCommand(IMailOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "返信元メールの Entry ID です。",
            Required = true,
        };
        var body = new Option<string>("--body")
        {
            Description = "返信本文です。",
            Required = true,
        };
        var command = new Command("reply", "メールへの返信を下書き保存します。");
        command.Options.Add(entryId);
        command.Options.Add(body);
        command.SetAction(result => Execute(result.GetValue(entryId), result.GetValue(body)));
        return command;
    }

    private int Execute(string? entryId, string? body)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id に返信元メールの Entry ID を指定してください。");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(body))
        {
            error.WriteLine("--body に返信本文を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "返信の下書き保存", () =>
        {
            var result = operations.Reply(entryId!, body!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                result.EntryId,
                result.DraftEntryId,
                result.Subject,
                saved = true,
            });
        });
    }
}
