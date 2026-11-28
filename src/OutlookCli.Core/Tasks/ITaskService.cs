namespace OutlookCli.Core.Tasks;

public interface ITaskService
{
    IEnumerable<TaskSummary> List();

    TaskSummary Add(string subject, DateTime? due);

    TaskChange Complete(string entryId);
}
