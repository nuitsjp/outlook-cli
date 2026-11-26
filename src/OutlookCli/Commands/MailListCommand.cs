using System.CommandLine;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailListCommand(IMailQuery query, TextWriter output, TextWriter error)
{
    public Command Create()
    {
        var unread = new Option<bool>("--unread") { Description = "未読メールだけを取得します。" };
        var command = new Command("list", "既定の受信トレイ直下のメールを受信日時の新しい順に取得します。");
        command.Options.Add(unread);
        command.SetAction(result => Execute(result.GetValue(unread)));
        return command;
    }

    private int Execute(bool unreadOnly)
    {
        try
        {
            foreach (var mail in query.List(unreadOnly))
            {
                MailJson.WriteSummary(output, mail);
            }

            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"受信トレイの取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
