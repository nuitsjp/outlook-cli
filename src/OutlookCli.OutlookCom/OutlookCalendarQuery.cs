using System.Globalization;
using System.Runtime.InteropServices;
using OutlookCli.Core.Calendar;

namespace OutlookCli.OutlookCom;

public sealed class OutlookCalendarQuery : ICalendarQuery, ICalendarOperations
{
    private const int CalendarFolder = 9; // OlDefaultFolders.olFolderCalendar
    private const int AppointmentItemClass = 26; // OlObjectClass.olAppointment
    private const int AppointmentItemType = 1; // OlItemType.olAppointmentItem
    private const int Busy = 2; // OlBusyStatus.olBusy
    private const int NotRecurring = 0; // OlRecurrenceState.olApptNotRecurring
    private const int Master = 1; // OlRecurrenceState.olApptMaster
    private const int Occurrence = 2; // OlRecurrenceState.olApptOccurrence
    private const int Exception = 3; // OlRecurrenceState.olApptException
    private const string ExceptionReplaceTimeSchema =
        "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/0x82280040";

    public IEnumerable<CalendarSummary> List(DateTime from, DateTime to, string? owner = null) =>
        ComCleanup.Enumerate(cleanup => ListCore(cleanup, from, to, owner));

    private static IEnumerable<CalendarSummary> ListCore(ComCleanup cleanup, DateTime from, DateTime to, string? owner)
    {
        using var calendar = OutlookFolderConnection.Open(cleanup, CalendarFolder, owner);
        foreach (var (item, _, _, state) in ReadAppointments(calendar, from, to))
        {
            var appointment = (dynamic)item;
            var entryId = (string)appointment.EntryID;
            // For an expanded item Outlook returns the series EntryID.
            var originalUtc = state switch
            {
                NotRecurring => (DateTime?)null,
                Occurrence => ReadUtc((DateTime)appointment.StartUTC),
                Exception => ReadExceptionOriginalUtc(cleanup, item),
                _ => throw new InvalidOperationException("Outlook の予定の繰り返し状態を判別できませんでした。"),
            };
            yield return ReadSummary(item, CalendarLogic.EncodeId(entryId, originalUtc));
        }
    }

    private static IEnumerable<(object Item, DateTime Start, DateTime End, int State)> ReadAppointments(
        OutlookFolderConnection calendar, DateTime from, DateTime to)
    {
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
                try
                {
                    var appointment = (dynamic)item;
                    if ((int)appointment.Class != AppointmentItemClass)
                    {
                        continue;
                    }

                    var state = (int)appointment.RecurrenceState;
                    if (state == Master)
                    {
                        continue;
                    }

                    var start = (DateTime)appointment.Start;
                    var end = (DateTime)appointment.End;
                    if (!CalendarLogic.Overlaps(start, end, from, to))
                    {
                        continue;
                    }

                    yield return (item, start, end, state);
                }
                finally
                {
                    calendar.Cleanup.Release(item);
                }
            }
        }
        finally
        {
            calendar.Cleanup.Release(restricted);
            calendar.Cleanup.Release(items);
        }
    }

    public CalendarDetail Get(string id, string? owner = null) => ComCleanup.Run(cleanup =>
    {
        var (entryId, originalUtc) = CalendarLogic.DecodeId(id);
        using var calendar = OutlookFolderConnection.Open(cleanup, CalendarFolder, owner);
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

            EnsureInCalendar(cleanup, item, calendar.Folder);
            var state = (int)((dynamic)item).RecurrenceState;
            if (originalUtc is null && state == NotRecurring)
            {
                return ReadDetail(cleanup, item, id);
            }

            if (originalUtc is not DateTime original || state != Master)
            {
                throw new InvalidOperationException("指定した ID に対応する予定が見つかりません。");
            }

            pattern = ((dynamic)item).GetRecurrencePattern();
            occurrence = FindException(cleanup, pattern!, original);
            // A moved occurrence must be read through Exceptions. GetOccurrence is for unchanged instances.
            occurrence ??= ((dynamic)pattern!).GetOccurrence(original.ToLocalTime());
            if (occurrence is null)
            {
                throw new InvalidOperationException("指定した ID に対応する予定の回が見つかりません。");
            }

            return ReadDetail(cleanup, occurrence, id);
        }
        finally
        {
            cleanup.Release(occurrence);
            cleanup.Release(pattern);
            cleanup.Release(item);
        }
    });

    public CalendarSummary Create(string subject, DateTimeOffset start, DateTimeOffset end) => ComCleanup.Run(cleanup =>
    {
        using var calendar = OutlookFolderConnection.Open(cleanup, CalendarFolder);
        object? item = null;
        try
        {
            item = ((dynamic)calendar.Application).CreateItem(AppointmentItemType);
            var appointment = (dynamic)item!;
            appointment.Subject = subject;
            appointment.StartUTC = start.UtcDateTime;
            appointment.EndUTC = end.UtcDateTime;
            appointment.BusyStatus = Busy;
            appointment.Save();
            return ReadSummary(item!, CalendarLogic.EncodeId((string)appointment.EntryID, null));
        }
        finally
        {
            cleanup.Release(item);
        }
    });

    public IEnumerable<CalendarAvailability> GetAvailability(IReadOnlyList<string> attendees, DateTime from, DateTime to) =>
        ComCleanup.Enumerate(cleanup => GetAvailabilityCore(cleanup, attendees, from, to));

    private static IEnumerable<CalendarAvailability> GetAvailabilityCore(ComCleanup cleanup, IReadOnlyList<string> attendees, DateTime from, DateTime to)
    {
        using var calendar = OutlookFolderConnection.Open(cleanup, CalendarFolder);
        var ownerAddress = ReadCalendarOwnerAddress(calendar);
        var slotCount = (int)((to - from).TotalMinutes / CalendarLogic.MinutesPerSlot);
        string? ownerSlots = null;
        foreach (var attendee in attendees)
        {
            if (string.Equals(attendee, ownerAddress, StringComparison.OrdinalIgnoreCase))
            {
                ownerSlots ??= ReadLocalSlots(calendar, from, to);
                yield return new CalendarAvailability(attendee, from, to, CalendarLogic.MinutesPerSlot, ownerSlots);
                continue;
            }

            var recipient = calendar.ResolveRecipient(attendee);
            string? slots;
            try
            {
                slots = (string?)((dynamic)recipient).FreeBusy(from, CalendarLogic.MinutesPerSlot, true);
            }
            catch (COMException exception)
            {
                throw new COMException($"Outlook が {attendee} の空き時間情報を取得できませんでした。{exception.Message}",
                    exception.HResult);
            }
            finally
            {
                cleanup.Release(recipient);
            }

            if (slots is null || slots.Length < slotCount)
            {
                throw new InvalidOperationException($"指定期間の空き時間情報を取得できませんでした。{attendee}");
            }

            yield return new CalendarAvailability(attendee, from, to, CalendarLogic.MinutesPerSlot, slots.Substring(0, slotCount));
        }
    }

    private static string? ReadCalendarOwnerAddress(OutlookFolderConnection calendar)
    {
        var storeId = (string)((dynamic)calendar.Folder).StoreID;
        object? accounts = null;
        try
        {
            accounts = ((dynamic)calendar.Session).Accounts;
            var collection = (dynamic)accounts!;
            var count = (int)collection.Count;
            for (var index = 1; index <= count; index++)
            {
                object? account = null;
                object? store = null;
                try
                {
                    account = collection.Item(index);
                    store = ((dynamic)account!).DeliveryStore;
                    if (store is not null && string.Equals(
                        (string)((dynamic)store).StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                    {
                        return (string?)((dynamic)account!).SmtpAddress;
                    }
                }
                finally
                {
                    calendar.Cleanup.Release(store);
                    calendar.Cleanup.Release(account);
                }
            }

            return null;
        }
        finally
        {
            calendar.Cleanup.Release(accounts);
        }
    }

    private static string ReadLocalSlots(OutlookFolderConnection calendar, DateTime from, DateTime to)
    {
        var appointments = ReadAppointments(calendar, from, to)
            .Select(appointment => new CalendarBusyInterval(
                appointment.Start,
                appointment.End,
                (int)((dynamic)appointment.Item).BusyStatus));
        return CalendarLogic.CalculateSlots(appointments, from, to);
    }

    internal static object? FindException(ComCleanup cleanup, object pattern, DateTime originalUtc)
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
                    if (ReadExceptionOriginalUtc(cleanup, appointment!) == originalUtc)
                    {
                        var match = appointment;
                        appointment = null;
                        return match;
                    }
                }
                finally
                {
                    cleanup.Release(appointment);
                    cleanup.Release(exception);
                }
            }

            return null;
        }
        finally
        {
            cleanup.Release(exceptions);
        }
    }

    private static DateTime ReadExceptionOriginalUtc(ComCleanup cleanup, object appointment)
    {
        object? accessor = null;
        try
        {
            accessor = ((dynamic)appointment).PropertyAccessor;
            return ReadUtc((DateTime)((dynamic)accessor!).GetProperty(ExceptionReplaceTimeSchema));
        }
        finally
        {
            cleanup.Release(accessor);
        }
    }

    private static void EnsureInCalendar(ComCleanup cleanup, object item, object folder)
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
                throw new InvalidOperationException("指定した予定は対象の予定表にありません。");
            }
        }
        finally
        {
            cleanup.Release(parent);
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

    private static CalendarDetail ReadDetail(ComCleanup cleanup, object item, string id)
    {
        var appointment = (dynamic)item;
        return new CalendarDetail(
            ReadSummary(item, id),
            (string?)appointment.Organizer ?? string.Empty,
            ReadAttendees(cleanup, item),
            (string?)appointment.Body ?? string.Empty);
    }

    private static IReadOnlyList<CalendarAttendee> ReadAttendees(ComCleanup cleanup, object item)
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
                    cleanup.Release(recipient);
                }
            }
        }
        finally
        {
            cleanup.Release(recipients);
        }

        return attendees;
    }

    private static DateTime ReadUtc(DateTime value) => DateTime.SpecifyKind(value, DateTimeKind.Utc);

}
