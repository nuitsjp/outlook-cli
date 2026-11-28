using System.CommandLine;
using OutlookCli.Core.Mail;
using OutlookCli.OutlookCom;

namespace OutlookCli.Commands;

internal sealed class MailSendCommand(IMailSender sender, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var to = new Option<string>("--to") { Description = "宛先のメールアドレスです。", Required = true };
        var subject = new Option<string>("--subject") { Description = "メールの件名です。", Required = true };
        var body = new Option<string>("--body") { Description = "プレーンテキストの本文です。", Required = true };
        var command = new Command("send", "既定のアカウントからメールを送信します。");
        command.Options.Add(to);
        command.Options.Add(subject);
        command.Options.Add(body);
        command.SetAction(result => Execute(
            result.GetRequiredValue(to), result.GetRequiredValue(subject), result.GetRequiredValue(body)));
        return command;
    }

    private int Execute(string to, string subject, string body)
    {
        if (string.IsNullOrWhiteSpace(to))
        {
            error.WriteLine("--to に宛先のメールアドレスを指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "Outlook への送信依頼", () =>
        {
            try
            {
                sender.Send(to, subject, body);
            }
            catch (OutlookComCleanupException exception) when (exception.MailSubmitted && exception.PrimaryException is null)
            {
                CommandExecution.ReportCleanupFailures(error, exception);
            }

            json.WriteLine(new { schemaVersion = 1, submitted = true });
        });
    }
}
