using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class CalendarLogicTests
{
    [Fact]
    public void Encode_and_decode_round_trip_a_non_recurring_id()
    {
        var encoded = CalendarLogic.EncodeId("0011AABB", null);

        var decoded = CalendarLogic.DecodeId(encoded);

        Assert.Equal("oc1_0011AABB", encoded);
        Assert.Equal("0011AABB", decoded.EntryId);
        Assert.Null(decoded.OriginalUtc);
    }

    [Fact]
    public void Encode_and_decode_round_trip_the_original_occurrence_time_as_utc()
    {
        var original = new DateTime(2026, 12, 1, 4, 30, 0, DateTimeKind.Utc);

        var decoded = CalendarLogic.DecodeId(CalendarLogic.EncodeId("ABCDEF", original));

        Assert.Equal("ABCDEF", decoded.EntryId);
        Assert.Equal(original, decoded.OriginalUtc);
        Assert.Equal(DateTimeKind.Utc, decoded.OriginalUtc!.Value.Kind);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("calendar-id")]
    [InlineData("oc1_")]
    [InlineData("oc1_GHI")]
    [InlineData("oc1_A.1.2")]
    [InlineData("oc1_A.not-a-number")]
    [InlineData("oc1_A.-1")]
    public void Decode_rejects_an_invalid_id(string? id)
    {
        Assert.Throws<InvalidOperationException>(() => CalendarLogic.DecodeId(id));
    }

    [Fact]
    public void Decode_rejects_ticks_beyond_the_datetime_range()
    {
        var id = $"oc1_A.{(long)DateTime.MaxValue.Ticks + 1}";

        Assert.Throws<InvalidOperationException>(() => CalendarLogic.DecodeId(id));
    }

    [Theory]
    [InlineData(10, 10, 10, 11, true)]
    [InlineData(11, 11, 10, 11, false)]
    [InlineData(11, 10, 10, 11, false)]
    [InlineData(11, 12, 10, 11, false)]
    [InlineData(9, 11, 10, 11, true)]
    [InlineData(10, 11, 10, 11, true)]
    [InlineData(9, 10, 10, 11, false)]
    public void Overlaps_applies_half_open_boundaries_and_zero_length_events(
        int startHour,
        int endHour,
        int fromHour,
        int toHour,
        bool expected)
    {
        var day = new DateTime(2026, 12, 1);

        var actual = CalendarLogic.Overlaps(
            day.AddHours(startHour),
            day.AddHours(endHour),
            day.AddHours(fromHour),
            day.AddHours(toHour));

        Assert.Equal(expected, actual);
    }

    [Fact]
    public void CalculateSlots_uses_thirty_minute_slots_and_the_highest_busy_priority()
    {
        var day = new DateTime(2026, 12, 1);
        var from = day.AddHours(10);
        var to = day.AddHours(12);
        var appointments = new[]
        {
            new CalendarBusyInterval(from, from.AddMinutes(30), 4),
            new CalendarBusyInterval(from.AddMinutes(15), from.AddMinutes(75), 1),
            new CalendarBusyInterval(from.AddMinutes(60), from.AddMinutes(90), 2),
            new CalendarBusyInterval(from.AddMinutes(90), to, 3),
            new CalendarBusyInterval(from, from.AddMinutes(30), 0),
        };

        var slots = CalendarLogic.CalculateSlots(appointments, from, to);

        Assert.Equal("1123", slots);
    }

    [Fact]
    public void CalculateSlots_clamps_partial_events_to_the_requested_window()
    {
        var day = new DateTime(2026, 12, 1);
        var from = day.AddHours(10);
        var to = day.AddHours(11);
        var appointments = new[]
        {
            new CalendarBusyInterval(from.AddMinutes(-15), from.AddMinutes(15), 2),
            new CalendarBusyInterval(from.AddMinutes(30), to.AddMinutes(30), 3),
        };

        var slots = CalendarLogic.CalculateSlots(appointments, from, to);

        Assert.Equal("23", slots);
    }

    [Fact]
    public void CalculateSlots_ignores_zero_length_events_even_when_they_overlap()
    {
        var day = new DateTime(2026, 12, 1);
        var from = day.AddHours(10);

        var slots = CalendarLogic.CalculateSlots(
            [new CalendarBusyInterval(from.AddMinutes(15), from.AddMinutes(15), 2)],
            from,
            from.AddHours(1));

        Assert.Equal("00", slots);
    }

    [Fact]
    public void CalculateSlots_rejects_an_unknown_busy_status()
    {
        var day = new DateTime(2026, 12, 1);

        Assert.Throws<InvalidOperationException>(() => CalendarLogic.CalculateSlots(
            [new CalendarBusyInterval(day.AddHours(10), day.AddHours(11), 99)],
            day.AddHours(10),
            day.AddHours(11)));
    }
}
