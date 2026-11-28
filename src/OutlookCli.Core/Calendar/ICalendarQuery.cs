namespace OutlookCli.Core.Calendar;

public interface ICalendarQuery
{
    IEnumerable<CalendarSummary> List(DateTime from, DateTime to, string? owner = null);

    CalendarDetail Get(string id, string? owner = null);
}
