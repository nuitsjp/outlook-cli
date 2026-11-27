using System.CommandLine;
using System.Runtime.InteropServices;
using OutlookCli.Core.Calendar;

namespace OutlookCli.Commands;

internal sealed class CalendarShowCommand(ICalendarQuery query, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var id = new Option<string>("--id")
        {
            Description = "一覧で取得した予定の ID です。",
            Required = true,
        };
        var command = new Command("show", "指定した予定の明細を取得します。");
        command.Options.Add(id);
        command.SetAction(result => Execute(result.GetValue(id)));
        return command;
    }

    private int Execute(string? id)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            error.WriteLine("--id に予定の ID を指定してください。");
            return 1;
        }

        try
        {
            var appointment = query.Get(id!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                appointment.Summary.Id,
                appointment.Summary.Subject,
                appointment.Summary.StartAt,
                appointment.Summary.EndAt,
                appointment.Summary.IsAllDay,
                appointment.Summary.Location,
                appointment.Organizer,
                appointment.Attendees,
                appointment.Body,
            });
            return 0;
        }
        catch (COMException exception)
        {
            error.WriteLine($"予定明細の取得に失敗しました (HRESULT 0x{exception.HResult:X8})。{exception.Message}");
            return 1;
        }
        catch (InvalidOperationException exception)
        {
            error.WriteLine(exception.Message);
            return 1;
        }
    }
}
