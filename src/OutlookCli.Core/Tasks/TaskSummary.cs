namespace OutlookCli.Core.Tasks;

public sealed record TaskSummary(string EntryId, string Subject, DateTime? Due)
{
    public DateTime? DueDate => Due;
}
