using System.CommandLine;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookCli.Core.Calendar;

namespace OutlookCli.Commands;

internal sealed class CalendarListCommand(ICalendarQuery query, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var from = new Option<string>("--from")
        {
            Description = "取得期間の開始日です (YYYY-MM-DD)。",
            Required = true,
        };
        var to = new Option<string>("--to")
        {
            Description = "取得期間の終了日です (YYYY-MM-DD、この日は含みません)。",
            Required = true,
        };
        var owner = new Option<string>("--owner") { Description = "共有予定表の所有者のメールアドレスです。" };
        var command = new Command("list", "指定期間の予定を取得します。");
        command.Options.Add(from);
        command.Options.Add(to);
        command.Options.Add(owner);
        command.SetAction(result => Execute(result.GetValue(from), result.GetValue(to), result.GetValue(owner)));
        return command;
    }

    private int Execute(string? from, string? to, string? owner)
    {
        if (!DateTime.TryParseExact(from, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var start)
            || !DateTime.TryParseExact(to, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var end))
        {
            error.WriteLine("--from と --to に YYYY-MM-DD 形式の日付を指定してください。");
            return 1;
        }

        if (start >= end)
        {
            error.WriteLine("--to には --from より後の日付を指定してください。");
            return 1;
        }

        if (owner is not null && string.IsNullOrWhiteSpace(owner))
        {
            error.WriteLine("--owner に共有予定表の所有者を指定してください。");
            return 1;
        }

        try
        {
            foreach (var appointment in query.List(start, end, owner))
            {
                json.WriteLine(new
                {
                    schemaVersion = 1,
                    appointment.Id,
                    appointment.Subject,
                    appointment.StartAt,
                    appointment.EndAt,
                    appointment.IsAllDay,
                    appointment.Location,
                });
            }

            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"予定一覧の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
