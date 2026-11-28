namespace OutlookCli.Core.Calendar;

public sealed record CalendarDetail(
    CalendarSummary Summary,
    string Organizer,
    IReadOnlyList<CalendarAttendee> Attendees,
    string Body);

public sealed record CalendarAttendee(string Name, string Type);
