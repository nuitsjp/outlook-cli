using System.CommandLine;
using System.Globalization;
using OutlookCli.Core.Calendar;

namespace OutlookCli.Commands;

internal sealed class CalendarCreateCommand(ICalendarOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var subject = new Option<string>("--subject") { Required = true, Description = "予定の件名です。" };
        var start = new Option<string>("--start") { Required = true, Description = "オフセット付きの開始日時です。" };
        var end = new Option<string>("--end") { Required = true, Description = "オフセット付きの終了日時です。" };
        var command = new Command("create", "既定の予定表に単発の予定を登録します。");
        command.Options.Add(subject);
        command.Options.Add(start);
        command.Options.Add(end);
        command.SetAction(result => Execute(result.GetValue(subject), result.GetValue(start), result.GetValue(end)));
        return command;
    }

    private int Execute(string? subject, string? start, string? end)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            error.WriteLine("--subject に予定の件名を指定してください。");
            return 1;
        }

        const string dateTimeFormat = "yyyy-MM-dd'T'HH:mm";
        string[] formats = [dateTimeFormat + "zzz", dateTimeFormat + "'Z'"];
        if (!DateTimeOffset.TryParseExact(start, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var startAt)
            || !DateTimeOffset.TryParseExact(end, formats, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out var endAt)
            || endAt <= startAt)
        {
            error.WriteLine("--start と --end にオフセット付きの日時（例: 2026-12-01T13:00+09:00）を指定し、--end を --start より後にしてください。");
            return 1;
        }

        return CommandExecution.Run(error, "予定の登録", () =>
        {
            var appointment = operations.Create(subject!, startAt, endAt);
            json.WriteLine(new
            {
                schemaVersion = 1,
                appointment.Id,
                appointment.Subject,
                appointment.StartAt,
                appointment.EndAt,
            });
        });
    }
}
