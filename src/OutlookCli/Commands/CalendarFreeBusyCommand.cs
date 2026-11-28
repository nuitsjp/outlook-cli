using System.CommandLine;
using System.Globalization;
using OutlookCli.Core.Calendar;

namespace OutlookCli.Commands;

internal sealed class CalendarFreeBusyCommand(ICalendarOperations operations, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var attendees = new Option<string[]>("--attendee")
        {
            Required = true,
            Arity = ArgumentArity.OneOrMore,
            Description = "参加者のメールアドレスです。複数人を指定できます。",
        };
        var from = new Option<string>("--from") { Required = true, Description = "取得期間の開始日です (YYYY-MM-DD)。" };
        var to = new Option<string>("--to") { Required = true, Description = "取得期間の終了日です (YYYY-MM-DD、この日は含みません)。" };
        var command = new Command("free-busy", "参加者の空き時間情報を取得します。");
        command.Options.Add(attendees);
        command.Options.Add(from);
        command.Options.Add(to);
        command.SetAction(result => Execute(result.GetValue(attendees), result.GetValue(from), result.GetValue(to)));
        return command;
    }

    private int Execute(string[]? attendees, string? from, string? to)
    {
        if (attendees is null || attendees.Length == 0 || attendees.Any(string.IsNullOrWhiteSpace))
        {
            error.WriteLine("--attendee に参加者のメールアドレスを指定してください。");
            return 1;
        }

        if (!DateTime.TryParseExact(from, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start)
            || !DateTime.TryParseExact(to, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end)
            || start >= end
            || (start.Year < 9999 || start.Month < 12) && end > start.AddMonths(1))
        {
            error.WriteLine("--from と --to に YYYY-MM-DD 形式の日付を指定してください。取得期間は1日から1か月です。");
            return 1;
        }

        return CommandExecution.RunList(error, "空き時間情報の取得", () => operations.GetAvailability(attendees, start, end),
            availability => json.WriteLine(new
            {
                schemaVersion = 1,
                availability.Attendee,
                From = availability.From.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                To = availability.To.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                availability.MinutesPerSlot,
                availability.Slots,
            }));
    }
}
