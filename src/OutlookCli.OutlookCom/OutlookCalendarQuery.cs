using System.Globalization;
using OutlookCli.Core.Calendar;

namespace OutlookCli.OutlookCom;

public sealed class OutlookCalendarQuery : ICalendarQuery
{
    private const int CalendarFolder = 9; // OlDefaultFolders.olFolderCalendar
    private const int AppointmentItemClass = 26; // OlObjectClass.olAppointment
    private const int NotRecurring = 0; // OlRecurrenceState.olApptNotRecurring
    private const int Master = 1; // OlRecurrenceState.olApptMaster
    private const int Occurrence = 2; // OlRecurrenceState.olApptOccurrence
    private const int Exception = 3; // OlRecurrenceState.olApptException
    private const string IdPrefix = "oc1_";
    private const string ExceptionReplaceTimeSchema =
        "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x82280040";

    public IEnumerable<CalendarSummary> List(DateTime from, DateTime to)
    {
        using var calendar = OutlookFolderConnection.Open(CalendarFolder);
        object? items = null;
        object? restricted = null;
        try
        {
            items = ((dynamic)calendar.Folder).Items;
            var appointments = (dynamic)items!;
            // Outlook expands recurring items when sorting precedes IncludeRecurrences and Restrict.
            appointments.Sort("[Start]", false);
            appointments.IncludeRecurrences = true;
            var fromText = from.ToString("g", CultureInfo.CurrentCulture);
            var toText = to.ToString("g", CultureInfo.CurrentCulture);
            restricted = appointments.Restrict($"[Start] < '{toText}' AND [End] >= '{fromText}'");

            var results = (dynamic)restricted!;
            for (object? item = results.GetFirst(); item is not null; item = results.GetNext())
            {
                CalendarSummary summary;
                try
                {
                    var appointment = (dynamic)item;
                    if ((int)appointment.Class != AppointmentItemClass
                        || (int)appointment.RecurrenceState == Master)
                    {
                        continue;
                    }

                    var start = (DateTime)appointment.Start;
                    var end = (DateTime)appointment.End;
                    var overlaps = start == end
                        ? start >= from && start < to
                        : start < to && end > from;
                    if (!overlaps)
                    {
                        continue;
                    }

                    var entryId = (string)appointment.EntryID;
                    // For an expanded item Outlook returns the series EntryID.
                    var originalUtc = (int)appointment.RecurrenceState switch
                    {
                        NotRecurring => (DateTime?)null,
                        Occurrence => ReadUtc((DateTime)appointment.StartUTC),
                        Exception => ReadExceptionOriginalUtc(item),
                        _ => throw new InvalidOperationException("Outlook の予定の繰り返し状態を判別できませんでした。"),
                    };
                    summary = ReadSummary(item, EncodeId(entryId, originalUtc));
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
            OutlookFolderConnection.Release(restricted);
            OutlookFolderConnection.Release(items);
        }
    }

    public CalendarDetail Get(string id)
    {
        var (entryId, originalUtc) = DecodeId(id);
        using var calendar = OutlookFolderConnection.Open(CalendarFolder);
        object? item = null;
        object? pattern = null;
        object? occurrence = null;
        try
        {
            item = calendar.GetItem(entryId);
            if (item is null || (int)((dynamic)item).Class != AppointmentItemClass)
            {
                throw new InvalidOperationException("指定した ID に対応する予定が見つかりません。");
            }

            EnsureInDefaultCalendar(item, calendar.Folder);
            var state = (int)((dynamic)item).RecurrenceState;
            if (originalUtc is null && state == NotRecurring)
            {
                return ReadDetail(item, id);
            }

            if (originalUtc is not DateTime original || state != Master)
            {
                throw new InvalidOperationException("指定した ID に対応する予定が見つかりません。");
            }

            pattern = ((dynamic)item).GetRecurrencePattern();
            occurrence = FindException(pattern!, original);
            // A moved occurrence must be read through Exceptions. GetOccurrence is for unchanged instances.
            occurrence ??= ((dynamic)pattern!).GetOccurrence(original.ToLocalTime());
            if (occurrence is null)
            {
                throw new InvalidOperationException("指定した ID に対応する予定の回が見つかりません。");
            }

            return ReadDetail(occurrence, id);
        }
        finally
        {
            OutlookFolderConnection.Release(occurrence);
            OutlookFolderConnection.Release(pattern);
            OutlookFolderConnection.Release(item);
        }
    }

    private static object? FindException(object pattern, DateTime originalUtc)
    {
        object? exceptions = null;
        try
        {
            exceptions = ((dynamic)pattern).Exceptions;
            var collection = (dynamic)exceptions!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? exception = null;
                object? appointment = null;
                try
                {
                    exception = collection.Item(index);
                    if ((bool)((dynamic)exception!).Deleted)
                    {
                        continue;
                    }

                    appointment = ((dynamic)exception!).AppointmentItem;
                    if (ReadExceptionOriginalUtc(appointment!) == originalUtc)
                    {
                        var match = appointment;
                        appointment = null;
                        return match;
                    }
                }
                finally
                {
                    OutlookFolderConnection.Release(appointment);
                    OutlookFolderConnection.Release(exception);
                }
            }

            return null;
        }
        finally
        {
            OutlookFolderConnection.Release(exceptions);
        }
    }

    private static DateTime ReadExceptionOriginalUtc(object appointment)
    {
        object? accessor = null;
        try
        {
            accessor = ((dynamic)appointment).PropertyAccessor;
            return ReadUtc((DateTime)((dynamic)accessor!).GetProperty(ExceptionReplaceTimeSchema));
        }
        finally
        {
            OutlookFolderConnection.Release(accessor);
        }
    }

    private static void EnsureInDefaultCalendar(object item, object folder)
    {
        object? parent = null;
        try
        {
            parent = ((dynamic)item).Parent;
            if (!string.Equals(
                (string)((dynamic)parent!).EntryID,
                (string)((dynamic)folder).EntryID,
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("指定した予定は既定の予定表にありません。");
            }
        }
        finally
        {
            OutlookFolderConnection.Release(parent);
        }
    }

    private static CalendarSummary ReadSummary(object item, string id)
    {
        var appointment = (dynamic)item;
        return new CalendarSummary(
            id,
            (string?)appointment.Subject ?? string.Empty,
            new DateTimeOffset(ReadUtc((DateTime)appointment.StartUTC)).ToLocalTime(),
            new DateTimeOffset(ReadUtc((DateTime)appointment.EndUTC)).ToLocalTime(),
            (bool)appointment.AllDayEvent,
            (string?)appointment.Location ?? string.Empty);
    }

    private static CalendarDetail ReadDetail(object item, string id)
    {
        var appointment = (dynamic)item;
        return new CalendarDetail(
            ReadSummary(item, id),
            (string?)appointment.Organizer ?? string.Empty,
            ReadAttendees(item),
            (string?)appointment.Body ?? string.Empty);
    }

    private static IReadOnlyList<CalendarAttendee> ReadAttendees(object item)
    {
        var attendees = new List<CalendarAttendee>();
        object? recipients = null;
        try
        {
            recipients = ((dynamic)item).Recipients;
            var collection = (dynamic)recipients!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? recipient = null;
                try
                {
                    recipient = collection.Item(index);
                    var type = (int)((dynamic)recipient!).Type switch
                    {
                        0 => null, // The organizer is returned separately.
                        1 => "required",
                        2 => "optional",
                        3 => "resource",
                        _ => throw new InvalidOperationException("Outlook の予定の参加者種別を判別できませんでした。"),
                    };
                    if (type is not null)
                    {
                        attendees.Add(new CalendarAttendee((string?)((dynamic)recipient!).Name ?? string.Empty, type));
                    }
                }
                finally
                {
                    OutlookFolderConnection.Release(recipient);
                }
            }
        }
        finally
        {
            OutlookFolderConnection.Release(recipients);
        }

        return attendees;
    }

    private static DateTime ReadUtc(DateTime value) => DateTime.SpecifyKind(value, DateTimeKind.Utc);

    private static string EncodeId(string entryId, DateTime? originalUtc) =>
        IdPrefix + entryId + (originalUtc is DateTime original
            ? "." + original.Ticks.ToString(CultureInfo.InvariantCulture)
            : string.Empty);

    private static (string EntryId, DateTime? OriginalUtc) DecodeId(string id)
    {
        const string invalidId = "予定の ID が不正です。一覧で取得した ID を指定してください。";
        if (!id.StartsWith(IdPrefix, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(invalidId);
        }

        var parts = id.Substring(IdPrefix.Length).Split('.');
        var entryId = parts[0];
        if (entryId.Length == 0 || !entryId.All(Uri.IsHexDigit))
        {
            throw new InvalidOperationException(invalidId);
        }

        if (parts.Length == 1)
        {
            return (entryId, null);
        }

        if (parts.Length != 2
            || !long.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var ticks)
            || ticks > DateTime.MaxValue.Ticks)
        {
            throw new InvalidOperationException(invalidId);
        }

        return (entryId, new DateTime(ticks, DateTimeKind.Utc));
    }
}
