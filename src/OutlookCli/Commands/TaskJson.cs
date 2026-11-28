using OutlookCli.Core.Tasks;

namespace OutlookCli.Commands;

internal static class TaskJson
{
    public static void WriteSummary(JsonOutput json, TaskSummary task)
    {
        json.WriteLine(new
        {
            schemaVersion = 1,
            task.EntryId,
            task.Subject,
            task.Due,
        });
    }
}
