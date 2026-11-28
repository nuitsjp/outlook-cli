using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class CalendarOccurrenceOwnershipTests
{
    [Fact]
    public void Matching_exception_transfers_appointment_to_caller_for_final_release()
    {
        var originalUtc = new DateTime(2026, 11, 28, 1, 2, 3, DateTimeKind.Utc);
        var appointment = new FakeAppointment(originalUtc);
        var exception = new FakeExceptionItem(appointment);
        var pattern = new FakePattern(exception);
        var released = new List<object>();
        var cleanup = new ComCleanup(released.Add);

        var selected = OutlookCalendarQuery.FindException(cleanup, pattern, originalUtc);

        Assert.Same(appointment, selected);
        Assert.Collection(released,
            value => Assert.Same(appointment.PropertyAccessor, value),
            value => Assert.Same(exception, value),
            value => Assert.Same(pattern.Exceptions, value));

        cleanup.Release(selected);

        Assert.Equal(4, released.Count);
        Assert.Same(appointment, released[3]);
        Assert.Equal(1, released.Count(value => ReferenceEquals(appointment, value)));
    }

    [Fact]
    public void Release_failure_for_previous_exception_does_not_cancel_later_transfer()
    {
        var originalUtc = new DateTime(2026, 11, 28, 1, 2, 3, DateTimeKind.Utc);
        var appointment = new FakeAppointment(originalUtc);
        var previousAppointment = new FakeAppointment(originalUtc.AddMinutes(-30));
        var previous = new FakeExceptionItem(previousAppointment);
        var selectedException = new FakeExceptionItem(appointment);
        var pattern = new FakePattern(previous, selectedException);
        var releaseFailure = new InvalidOperationException("previous exception release");
        var released = new List<object>();
        var cleanup = new ComCleanup(value =>
        {
            released.Add(value);
            if (ReferenceEquals(previous, value))
            {
                throw releaseFailure;
            }
        });

        var selected = OutlookCalendarQuery.FindException(cleanup, pattern, originalUtc);

        Assert.Same(appointment, selected);
        Assert.Same(releaseFailure, Assert.Single(cleanup.Failures));
        Assert.Collection(released,
            value => Assert.Same(previousAppointment.PropertyAccessor, value),
            value => Assert.Same(previousAppointment, value),
            value => Assert.Same(previous, value),
            value => Assert.Same(appointment.PropertyAccessor, value),
            value => Assert.Same(selectedException, value),
            value => Assert.Same(pattern.Exceptions, value));

        cleanup.Release(selected);

        Assert.Equal(7, released.Count);
        Assert.Same(appointment, released[6]);
        Assert.Equal(1, released.Count(value => ReferenceEquals(appointment, value)));
    }

    public sealed class FakePattern(params FakeExceptionItem[] items)
    {
        public FakeExceptionCollection Exceptions { get; } = new(items);
    }

    public sealed class FakeExceptionCollection(FakeExceptionItem[] items)
    {
        public int Count => items.Length;

        public FakeExceptionItem Item(int index) => items[index - 1];
    }

    public sealed class FakeExceptionItem(FakeAppointment appointment)
    {
        public bool Deleted => false;

        public FakeAppointment AppointmentItem { get; } = appointment;
    }

    public sealed class FakeAppointment(DateTime originalUtc)
    {
        public FakePropertyAccessor PropertyAccessor { get; } = new(originalUtc);
    }

    public sealed class FakePropertyAccessor(DateTime originalUtc)
    {
        public DateTime GetProperty(string _) => originalUtc;
    }
}
