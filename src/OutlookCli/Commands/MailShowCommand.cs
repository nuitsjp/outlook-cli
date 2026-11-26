using System.CommandLine;
using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailShowCommand(IMailQuery query, TextWriter output, TextWriter error)
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

        try
        {
            var mail = query.Get(entryId);
            output.WriteLine(JsonSerializer.Serialize(new
            {
                schemaVersion = 1,
                mail.Summary.EntryId,
                mail.Summary.ReceivedAt,
                mail.Summary.SenderName,
                mail.Summary.Subject,
                mail.To,
                mail.Cc,
                mail.Bcc,
                mail.Body,
            }, MailJson.Options));
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"メール明細の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
