using System.CommandLine;
using System.Globalization;
using OutlookCli.Core.Tasks;

namespace OutlookCli.Commands;

internal sealed class TaskCommand(ITaskService service, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("task", "Outlook のタスクを操作します。");
        command.Subcommands.Add(CreateList());
        command.Subcommands.Add(CreateAdd());
        command.Subcommands.Add(CreateComplete());
        return command;
    }

    private Command CreateList()
    {
        var command = new Command("list", "未完了のタスクを取得します。");
        command.SetAction(_ => ExecuteList());
        return command;
    }

    private Command CreateAdd()
    {
        var subject = new Option<string>("--subject")
        {
            Description = "タスクの件名です。",
            Required = true,
        };
        var due = new Option<string>("--due")
        {
            Description = "タスクの期限です (YYYY-MM-DD)。",
        };
        var command = new Command("add", "タスクを追加します。");
        command.Options.Add(subject);
        command.Options.Add(due);
        command.SetAction(result => ExecuteAdd(
            result.GetRequiredValue(subject), result.GetValue(due)));
        return command;
    }

    private Command CreateComplete()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "完了するタスクの Entry ID です。",
            Required = true,
        };
        var command = new Command("complete", "指定したタスクを完了にします。");
        command.Options.Add(entryId);
        command.SetAction(result => ExecuteComplete(result.GetValue(entryId)));
        return command;
    }

    private int ExecuteList()
    {
        return CommandExecution.RunList(error, "タスク一覧の取得", service.List,
            task => TaskJson.WriteSummary(json, task));
    }

    private int ExecuteAdd(string subject, string? due)
    {
        if (string.IsNullOrWhiteSpace(subject))
        {
            error.WriteLine("--subject にタスクの件名を指定してください。");
            return 1;
        }

        DateTime? dueDate = null;
        if (due is not null)
        {
            if (!DateTime.TryParseExact(
                    due,
                    "yyyy-MM-dd",
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out var parsedDue))
            {
                error.WriteLine("--due に YYYY-MM-DD 形式の期限を指定してください。");
                return 1;
            }

            dueDate = parsedDue;
        }

        return CommandExecution.Run(error, "タスクの追加", () =>
            TaskJson.WriteSummary(json, service.Add(subject, dueDate)));
    }

    private int ExecuteComplete(string? entryId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id にタスクの Entry ID を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "タスクの完了", () =>
        {
            var change = service.Complete(entryId!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                change.EntryId,
                change.Changed,
            });
        });
    }
}
