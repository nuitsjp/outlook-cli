using System.CommandLine;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailWatchCommand(IMailWatcher watcher, TextWriter output, TextWriter error)
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
            watcher.Watch(
                mail =>
                {
                    MailJson.WriteSummary(output, mail);
                    output.Flush();
                },
                () => error.WriteLine("Outlook の受信イベントを購読しました。終了するには Ctrl+C を押してください。"),
                stopping.Token);
            error.WriteLine("Outlook の受信イベントの購読を解除しました。");
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"Outlook の受信監視に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
        finally
        {
            Console.CancelKeyPress -= Cancel;
        }
    }
}
