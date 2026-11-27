namespace OutlookCli.Core.Calendar;

public interface ICalendarQuery
{
    IEnumerable<CalendarSummary> List(DateTime from, DateTime to);

    CalendarDetail Get(string id);
}
