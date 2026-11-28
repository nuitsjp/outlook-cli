using OutlookCli.Core.Tasks;

namespace OutlookCli.OutlookCom;

public sealed class OutlookTaskService : ITaskService
{
    private const int TasksFolder = 13; // OlDefaultFolders.olFolderTasks
    private const int TaskItemType = 3; // OlItemType.olTaskItem
    private const int TaskItemClass = 48; // OlObjectClass.olTask

    public IEnumerable<TaskSummary> List()
    {
        using var tasks = OutlookFolderConnection.Open(TasksFolder);
        object? items = null;
        object? restrictedItems = null;

        try
        {
            items = ((dynamic)tasks.Folder).Items;
            restrictedItems = ((dynamic)items!).Restrict("[Complete] = False");
            var results = (dynamic)(restrictedItems ?? items!);
            results.Sort("[DueDate]", false);

            for (object? item = results.GetFirst(); item is not null; item = results.GetNext())
            {
                TaskSummary summary;
                try
                {
                    if (!IsTask(item))
                    {
                        continue;
                    }

                    summary = ReadSummary(item);
                }
                finally
                {
                    OutlookFolderConnection.Release(item);
                }

                yield return summary;
            }
        }
        finally
        {
            OutlookFolderConnection.Release(restrictedItems);
            OutlookFolderConnection.Release(items);
        }
    }

    public TaskSummary Add(string subject, DateTime? due)
    {
        using var tasks = OutlookFolderConnection.Open(TasksFolder);
        object? item = null;

        try
        {
            item = ((dynamic)tasks.Application).CreateItem(TaskItemType);
            var task = (dynamic)item;
            task.Subject = subject;
            if (due.HasValue)
            {
                task.DueDate = due.Value;
            }

            task.Save();
            return ReadSummary(item);
        }
        finally
        {
            OutlookFolderConnection.Release(item);
        }
    }

    public TaskChange Complete(string entryId)
    {
        using var tasks = OutlookFolderConnection.Open(TasksFolder);
        var item = tasks.GetItem(entryId);
        object? parent = null;
        try
        {
            if (!IsTask(item))
            {
                throw new InvalidOperationException("指定された ID に対応する未完了タスクが見つかりません。");
            }

            parent = ((dynamic)item!).Parent;
            if (parent is null
                || !string.Equals(
                    (string?)((dynamic)parent).EntryID,
                    (string?)((dynamic)tasks.Folder).EntryID,
                    StringComparison.Ordinal))
            {
                throw new InvalidOperationException("指定された ID は既定のタスクフォルダーにありません。");
            }

            var task = (dynamic)item!;
            var change = new TaskChange((string)task.EntryID, !(bool)task.Complete);
            if (change.Changed)
            {
                task.Complete = true;
                task.Save();
            }

            return change;
        }
        finally
        {
            OutlookFolderConnection.Release(parent);
            OutlookFolderConnection.Release(item);
        }
    }

    private static bool IsTask(object? item) =>
        item is not null && (int)((dynamic)item).Class == TaskItemClass;

    private static TaskSummary ReadSummary(object item)
    {
        var task = (dynamic)item;
        var due = ReadDue(task.DueDate);
        return new TaskSummary(
            (string)task.EntryID,
            (string?)task.Subject ?? string.Empty,
            due);
    }

    private static DateTime? ReadDue(object? value)
    {
        if (value is not DateTime due || due.Date == new DateTime(4501, 1, 1))
        {
            return null;
        }

        return DateTime.SpecifyKind(due.Date, DateTimeKind.Unspecified);
    }
}
