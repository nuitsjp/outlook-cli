namespace OutlookCli.Core.Calendar;

public interface ICalendarOperations
{
    CalendarSummary Create(string subject, DateTimeOffset start, DateTimeOffset end);

    IEnumerable<CalendarAvailability> GetAvailability(IReadOnlyList<string> attendees, DateTime from, DateTime to);
}

public sealed record CalendarAvailability(string Attendee, DateTime From, DateTime To, int MinutesPerSlot, string Slots);
