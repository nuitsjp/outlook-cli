using System.Runtime.InteropServices;
using OutlookCli.OutlookCom;
using Xunit;
using Xunit.Sdk;

namespace OutlookCli.Tests;

public sealed class ComCleanupTests
{
    [Fact]
    public void Release_continues_after_a_failure_without_shared_test_state()
    {
        var first = new object();
        var second = new object();
        var released = new List<object>();
        var cleanup = new ComCleanup(value =>
        {
            if (ReferenceEquals(value, first))
            {
                throw new InvalidOperationException("first release");
            }

            released.Add(value);
        });

        cleanup.Release(first);
        cleanup.Release(second);

        Assert.Single(cleanup.Failures);
        Assert.Equal([second], released);
    }

    [Fact]
    public void Run_returns_result_when_operation_and_cleanup_succeed()
    {
        var result = ComCleanup.Run(_ => 42);

        Assert.Equal(42, result);
    }

    [Fact]
    public void Run_rethrows_primary_exception_when_cleanup_is_clean()
    {
        var primary = new InvalidOperationException("operation");

        var exception = Assert.Throws<InvalidOperationException>(() =>
            ComCleanup.Run(_ => throw primary));

        Assert.Same(primary, exception);
    }

    [Fact]
    public void Release_records_invalid_values_and_continues_with_later_com_values()
    {
        var dictionary = CreateDictionary();
        try
        {
            var cleanup = new ComCleanup();

            cleanup.Release(new object());
            cleanup.Release(dictionary);

            Assert.Single(cleanup.Failures);
            Assert.Throws<InvalidComObjectException>(() => { _ = (int)((dynamic)dictionary).Count; });
        }
        finally
        {
            try
            {
                Marshal.FinalReleaseComObject(dictionary);
            }
            catch (InvalidComObjectException)
            {
            }
        }
    }

    [Fact]
    public void Run_preserves_primary_exception_and_collects_all_cleanup_failures()
    {
        var primary = new COMException("operation", unchecked((int)0x80004005));

        var exception = Assert.Throws<OutlookComCleanupException>(() =>
            ComCleanup.Run(cleanup =>
            {
                cleanup.Release(new object());
                cleanup.Release(new object());
                throw primary;
            }));

        Assert.Same(primary, exception.PrimaryException);
        Assert.Equal(primary.HResult, exception.HResult);
        Assert.Equal(2, exception.CleanupExceptions.Count);
        Assert.False(exception.MailSubmitted);
    }

    [Fact]
    public void Run_marks_mail_as_submitted_when_cleanup_fails_after_send()
    {
        var exception = Assert.Throws<OutlookComCleanupException>(() =>
            ComCleanup.Run(cleanup =>
            {
                cleanup.MarkMailSubmitted();
                cleanup.Release(new object());
            }));

        Assert.Null(exception.PrimaryException);
        Assert.True(exception.MailSubmitted);
        Assert.Single(exception.CleanupExceptions);
    }

    [Fact]
    public void Run_flattens_nested_cleanup_diagnostics_without_duplicate_primary()
    {
        var nested = new OutlookComCleanupException(
            null,
            [new InvalidOperationException("inner cleanup")]);

        var exception = Assert.Throws<OutlookComCleanupException>(() =>
            ComCleanup.Run(cleanup =>
            {
                cleanup.Release(new object());
                throw nested;
            }));

        Assert.Null(exception.PrimaryException);
        Assert.Equal(2, exception.CleanupExceptions.Count);
        Assert.DoesNotContain(nested, exception.CleanupExceptions);
    }

    [Fact]
    public void Enumerate_preserves_source_exception_and_cleanup_failures()
    {
        var primary = new InvalidOperationException("source");

        var exception = Assert.Throws<OutlookComCleanupException>(() =>
            ComCleanup.Enumerate<int>(cleanup => Source(cleanup, primary)).ToArray());

        Assert.Same(primary, exception.PrimaryException);
        Assert.Equal(2, exception.CleanupExceptions.Count);
    }

    [Fact]
    public void Enumerate_runs_cleanup_when_consumer_stops_early()
    {
        var enumerator = ComCleanup.Enumerate<int>(SourceWithEarlyStopCleanup).GetEnumerator();
        try
        {
            Assert.True(enumerator.MoveNext());

            var exception = Assert.Throws<OutlookComCleanupException>(() => enumerator.Dispose());

            Assert.Null(exception.PrimaryException);
            Assert.Single(exception.CleanupExceptions);
        }
        finally
        {
            enumerator.Dispose();
        }
    }

    [Fact]
    public void Enumeration_stops_on_release_failure_before_returning_another_row()
    {
        static IEnumerable<int> Rows(ComCleanup cleanup)
        {
            yield return 1;
            cleanup.Attempt(() => throw new InvalidOperationException("releasing an item"));
            yield return 2;
        }

        using var iterator = ComCleanup.Enumerate(Rows).GetEnumerator();
        Assert.True(iterator.MoveNext());
        Assert.Equal(1, iterator.Current);
        var error = Assert.Throws<OutlookComCleanupException>(() => iterator.MoveNext());
        Assert.Equal("releasing an item", Assert.Single(error.CleanupExceptions).Message);
    }

    private static IEnumerable<int> Source(ComCleanup cleanup, Exception primary)
    {
        try
        {
            yield return 1;
            cleanup.Release(new object());
            throw primary;
        }
        finally
        {
            cleanup.Release(new object());
        }
    }

    private static IEnumerable<int> SourceWithEarlyStopCleanup(ComCleanup cleanup)
    {
        try
        {
            yield return 1;
        }
        finally
        {
            cleanup.Release(new object());
        }
    }

    private static object CreateDictionary()
    {
        var type = Type.GetTypeFromProgID("Scripting.Dictionary");
        if (type is null)
        {
            throw SkipException.ForSkip("Scripting.Dictionary の COM 登録がありません。");
        }

        return Activator.CreateInstance(type)!;
    }
}
