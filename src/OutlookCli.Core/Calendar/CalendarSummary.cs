namespace OutlookCli.Core.Calendar;

public sealed record CalendarSummary(
    string Id,
    string Subject,
    DateTimeOffset StartAt,
    DateTimeOffset EndAt,
    bool IsAllDay,
    string Location);
