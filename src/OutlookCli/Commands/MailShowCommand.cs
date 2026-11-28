using System.CommandLine;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailShowCommand(IMailQuery query, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "一覧で取得したメールの Entry ID です。",
            Required = true,
        };
        var command = new Command("show", "指定したメールの明細を取得します。");
        command.Options.Add(entryId);
        command.SetAction(result => Execute(result.GetValue(entryId)));
        return command;
    }

    private int Execute(string? entryId)
    {
        if (entryId is null || string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にメールの Entry ID を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "メール明細の取得", () =>
            json.WriteLine(new MailDetailJson(query.Get(entryId))));
    }
}
