using System.CommandLine;
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
        var owner = new Option<string>("--owner") { Description = "共有予定表の所有者のメールアドレスです。" };
        var command = new Command("show", "指定した予定の明細を取得します。");
        command.Options.Add(id);
        command.Options.Add(owner);
        command.SetAction(result => Execute(result.GetValue(id), result.GetValue(owner)));
        return command;
    }

    private int Execute(string? id, string? owner)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            error.WriteLine("--id に予定の ID を指定してください。");
            return 1;
        }

        if (owner is not null && string.IsNullOrWhiteSpace(owner))
        {
            error.WriteLine("--owner に共有予定表の所有者を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "予定明細の取得", () =>
            json.WriteLine(new CalendarDetailJson(query.Get(id!, owner))));
    }
}
