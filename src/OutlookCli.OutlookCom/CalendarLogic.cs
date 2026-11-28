namespace OutlookCli.OutlookCom;

internal readonly record struct CalendarBusyInterval(DateTime Start, DateTime End, int BusyStatus);

internal static class CalendarLogic
{
    private const string IdPrefix = "oc1_";
    internal const int MinutesPerSlot = 30;

    internal static string EncodeId(string entryId, DateTime? originalUtc)
    {
        if (entryId is null)
        {
            throw new ArgumentNullException(nameof(entryId));
        }

        return IdPrefix + entryId + (originalUtc is DateTime original
            ? "." + original.Ticks.ToString(System.Globalization.CultureInfo.InvariantCulture)
            : string.Empty);
    }

    internal static (string EntryId, DateTime? OriginalUtc) DecodeId(string? id)
    {
        const string invalidId = "予定の ID が不正です。一覧で取得した ID を指定してください。";
        if (id is null || id.Length == 0 || !id.StartsWith(IdPrefix, StringComparison.Ordinal))
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
            || !long.TryParse(
                parts[1],
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out var ticks)
            || ticks < 0
            || ticks > DateTime.MaxValue.Ticks)
        {
            throw new InvalidOperationException(invalidId);
        }

        return (entryId, new DateTime(ticks, DateTimeKind.Utc));
    }

    internal static bool Overlaps(DateTime start, DateTime end, DateTime from, DateTime to) =>
        start == end
            ? start >= from && start < to
            : start < to && end > from;

    internal static string CalculateSlots(
        IEnumerable<CalendarBusyInterval> appointments,
        DateTime from,
        DateTime to)
    {
        if (appointments is null)
        {
            throw new ArgumentNullException(nameof(appointments));
        }

        var slotCount = (int)((to - from).TotalMinutes / MinutesPerSlot);
        var slots = new int[slotCount];
        foreach (var appointment in appointments)
        {
            var priority = ToPriority(appointment.BusyStatus);
            if (appointment.Start >= appointment.End
                || priority == 0
                || !Overlaps(appointment.Start, appointment.End, from, to))
            {
                continue;
            }

            var first = Math.Max(
                0,
                (int)Math.Floor((appointment.Start - from).TotalMinutes / MinutesPerSlot));
            var last = Math.Min(
                slotCount,
                (int)Math.Ceiling((appointment.End - from).TotalMinutes / MinutesPerSlot));
            for (var index = first; index < last; index++)
            {
                slots[index] = Math.Max(slots[index], priority);
            }
        }

        const string codesInPriorityOrder = "05123";
        return new string(Array.ConvertAll(slots, priority => codesInPriorityOrder[priority]));
    }

    private static int ToPriority(int busyStatus) => busyStatus switch
    {
        0 => 0, // olFree
        1 => 2, // olTentative
        2 => 3, // olBusy
        3 => 4, // olOutOfOffice
        4 => 1, // olWorkingElsewhere
        _ => throw new InvalidOperationException("Outlook の予定の公開方法を判別できませんでした。"),
    };
}
