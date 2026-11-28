using System.CommandLine;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailWatchCommand(IMailWatcher watcher, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("watch", "Outlook の新着メールを受信するたびに表示します。");
        command.SetAction(_ => Execute());
        return command;
    }

    private int Execute()
    {
        using var stopping = new CancellationTokenSource();
        void Cancel(object? sender, ConsoleCancelEventArgs args)
        {
            args.Cancel = true;
            stopping.Cancel();
        }

        Console.CancelKeyPress += Cancel;
        try
        {
            return CommandExecution.Run(error, "Outlook の受信監視", () =>
            {
                watcher.Watch(
                    mail =>
                    {
                        json.WriteLine(new MailSummaryJson(mail));
                        json.Flush();
                    },
                    () => error.WriteLine("Outlook の受信イベントを購読しました。終了するには Ctrl+C を押してください。"),
                    stopping.Token);
                error.WriteLine("Outlook の受信イベントの購読を解除しました。");
            });
        }
        finally
        {
            Console.CancelKeyPress -= Cancel;
        }
    }
}
