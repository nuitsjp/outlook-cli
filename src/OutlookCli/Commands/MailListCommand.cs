using System.CommandLine;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal sealed class MailListCommand(IMailQuery query, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var unread = new Option<bool>("--unread") { Description = "未読メールだけを取得します。" };
        var folder = new Option<string>("--folder")
        {
            Description = "取得対象メールのフォルダー Entry ID です。",
        };
        var subject = new Option<string>("--subject")
        {
            Description = "件名に含まれる文字列で絞り込みます。",
        };
        var since = new Option<string>("--since")
        {
            Description = "受信日時の下限を指定します。 (ISO 8601)",
        };
        var command = new Command("list", "メールを受信日時の新しい順に取得します。");
        command.Options.Add(unread);
        command.Options.Add(folder);
        command.Options.Add(subject);
        command.Options.Add(since);
        command.SetAction(result => Execute(
            result.GetValue(unread),
            result.GetValue(folder),
            result.GetValue(subject),
            result.GetValue(since)));
        return command;
    }

    private int Execute(bool unreadOnly, string? folder, string? subject, string? sinceText)
    {
        if (folder is not null && string.IsNullOrWhiteSpace(folder))
        {
            error.WriteLine("--folder にフォルダー Entry ID を指定してください。");
            return 1;
        }

        if (subject is not null && string.IsNullOrWhiteSpace(subject))
        {
            error.WriteLine("--subject に件名の検索文字列を指定してください。");
            return 1;
        }

        DateTimeOffset? since = null;
        if (sinceText is not null)
        {
            string[] formats = ["yyyy-MM-dd", "yyyy-MM-dd'T'HH:mmK", "yyyy-MM-dd'T'HH:mm:ssK"];
            if (!DateTimeOffset.TryParseExact(
                sinceText, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsedSince))
            {
                error.WriteLine("--since に ISO 8601 形式の日時 (例: 2026-11-25T00:00:00+09:00) を指定してください。");
                return 1;
            }

            since = parsedSince;
        }

        try
        {
            foreach (var mail in query.List(new MailListFilter(unreadOnly, folder, subject, since)))
            {
                MailJson.WriteSummary(json, mail);
            }

            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"メール一覧の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
