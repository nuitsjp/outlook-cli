using System.CommandLine;
using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailSendCommand(IMailSender sender, TextWriter output, TextWriter error)
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

        try
        {
            sender.Send(to, subject, body);
            output.WriteLine(JsonSerializer.Serialize(new { schemaVersion = 1, submitted = true }, MailJson.Options));
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"Outlook への送信依頼に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
