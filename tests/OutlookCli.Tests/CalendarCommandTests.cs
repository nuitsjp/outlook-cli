using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Calendar;
using OutlookCli.Core.Mail;
using Xunit;

namespace OutlookCli.Tests;

public sealed class CalendarCommandTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void List_passes_dates_as_unspecified_and_writes_one_json_line_per_item(bool asciiOnly)
    {
        var from = new DateTime(2026, 11, 1);
        var to = new DateTime(2026, 12, 1);
        var first = new CalendarSummary(
            "calendar-1",
            "\u4f1a\u8b70\u30fb\u30c6\u30b9\u30c8",
            new DateTimeOffset(2026, 11, 3, 9, 30, 0, TimeSpan.FromHours(9)),
            new DateTimeOffset(2026, 11, 3, 10, 0, 0, TimeSpan.FromHours(9)),
            false,
            "\u6771\u4eac\u30aa\u30d5\u30a3\u30b9");
        var second = new CalendarSummary(
            "calendar-2",
            "Second event",
            new DateTimeOffset(2026, 11, 10, 0, 0, 0, TimeSpan.Zero),
            new DateTimeOffset(2026, 11, 11, 0, 0, 0, TimeSpan.Zero),
            true,
            "");
        var query = new RecordingCalendarQuery(_ => [first, second]);
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly
            ? ["--ascii", "calendar", "list", "--from", "2026-11-01", "--to", "2026-12-01"]
            : ["calendar", "list", "--from", "2026-11-01", "--to", "2026-12-01"];
        var exitCode = OutlookCli.Program.Run(
            args, UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        var listCall = Assert.Single(query.ListArguments);
        Assert.Equal(from, listCall.From);
        Assert.Equal(DateTimeKind.Unspecified, listCall.From.Kind);
        Assert.Equal(to, listCall.To);
        Assert.Equal(DateTimeKind.Unspecified, listCall.To.Kind);

        var lines = ReadLines(output.ToString());
        Assert.Equal(2, lines.Count);
        AssertJsonEncoding(lines[0], first.Subject, asciiOnly);
        AssertJsonEncoding(lines[0], first.Location, asciiOnly);

        AssertSummaryJson(lines[0], first);
        AssertSummaryJson(lines[1], second);
    }

    [Fact]
    public void List_with_no_items_writes_no_stdout()
    {
        var query = new RecordingCalendarQuery(_ => Array.Empty<CalendarSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "list", "--from", "2026-11-01", "--to", "2026-12-01"],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(output.ToString());
        Assert.Empty(error.ToString());
        Assert.Single(query.ListArguments);
    }

    [Theory]
    [InlineData("2026-11-01", "2026-11-01")]
    [InlineData("2026-12-02", "2026-12-01")]
    [InlineData("2026/11/01", "2026-12-01")]
    [InlineData("2026-2-01", "2026-12-01")]
    public void List_rejects_invalid_date_range(string from, string to)
    {
        var query = new RecordingCalendarQuery(_ => throw new InvalidOperationException("List must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "list", "--from", from, "--to", to],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.ListArguments);
    }

    [Fact]
    public void List_without_from_returns_one_and_does_not_call_query()
    {
        var query = new RecordingCalendarQuery(_ => throw new InvalidOperationException("List must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "list", "--to", "2026-12-01"],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.ListArguments);
    }

    [Fact]
    public void List_without_to_returns_one_and_does_not_call_query()
    {
        var query = new RecordingCalendarQuery(_ => throw new InvalidOperationException("List must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "list", "--from", "2026-11-01"],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.ListArguments);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Show_calls_get_once_preserves_id_and_writes_summary_and_details(bool asciiOnly)
    {
        const string id = "  calendar-id-42  ";
        var summary = new CalendarSummary(
            "summary-id",
            "\u4f1a\u8b70\u306e\u8a73\u7d30",
            new DateTimeOffset(2026, 11, 25, 13, 0, 0, TimeSpan.FromHours(9)),
            new DateTimeOffset(2026, 11, 25, 14, 30, 0, TimeSpan.FromHours(9)),
            false,
            "Room A");
        var detail = new CalendarDetail(
            summary,
            "\u4e3b\u50ac\u8005",
            [
                new CalendarAttendee("Alice", "required"),
                new CalendarAttendee("\u4efb\u610f\u53c2\u52a0\u8005", "optional"),
                new CalendarAttendee("Projector", "resource"),
            ],
            "\u672c\u6587\u306e1\u884c\u76ee\n\u672c\u6587\u306e2\u884c\u76ee");
        var query = new RecordingCalendarQuery(_ => Array.Empty<CalendarSummary>(), _ => detail);
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly
            ? ["--ascii", "calendar", "show", "--id", id]
            : ["calendar", "show", "--id", id];
        var exitCode = OutlookCli.Program.Run(
            args, UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal([id], query.GetArguments);

        var line = Assert.Single(ReadLines(output.ToString()));
        AssertJsonEncoding(line, detail.Organizer, asciiOnly);
        AssertJsonEncoding(line, detail.Body.Split('\n')[0], asciiOnly);
        using var json = JsonDocument.Parse(line);
        var root = json.RootElement;
        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(summary.Id, root.GetProperty("id").GetString());
        Assert.Equal(summary.Subject, root.GetProperty("subject").GetString());
        Assert.Equal(summary.StartAt, root.GetProperty("startAt").GetDateTimeOffset());
        Assert.Equal(summary.EndAt, root.GetProperty("endAt").GetDateTimeOffset());
        Assert.Equal(summary.IsAllDay, root.GetProperty("isAllDay").GetBoolean());
        Assert.Equal(summary.Location, root.GetProperty("location").GetString());
        Assert.Equal(detail.Organizer, root.GetProperty("organizer").GetString());
        Assert.Equal(detail.Body, root.GetProperty("body").GetString());

        var attendees = root.GetProperty("attendees").EnumerateArray().ToArray();
        Assert.Equal(3, attendees.Length);
        Assert.Equal("Alice", attendees[0].GetProperty("name").GetString());
        Assert.Equal("required", attendees[0].GetProperty("type").GetString());
        Assert.Equal("\u4efb\u610f\u53c2\u52a0\u8005", attendees[1].GetProperty("name").GetString());
        Assert.Equal("optional", attendees[1].GetProperty("type").GetString());
        Assert.Equal("Projector", attendees[2].GetProperty("name").GetString());
        Assert.Equal("resource", attendees[2].GetProperty("type").GetString());
    }

    [Fact]
    public void Show_without_id_returns_one_and_does_not_call_query()
    {
        var query = new RecordingCalendarQuery(_ => Array.Empty<CalendarSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "show"],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.GetArguments);
    }

    [Fact]
    public void Show_with_whitespace_id_returns_one_and_does_not_call_query()
    {
        var query = new RecordingCalendarQuery(_ => Array.Empty<CalendarSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "show", "--id", "   "],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.GetArguments);
    }

    [Fact]
    public void Show_query_failure_returns_one_and_writes_diagnostic()
    {
        const string id = "calendar-id";
        var query = new RecordingCalendarQuery(
            _ => Array.Empty<CalendarSummary>(),
            _ => throw new COMException("Calendar failed", unchecked((int)0x80004005)));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = OutlookCli.Program.Run(
            ["calendar", "show", "--id", id],
            UnusedMailQuery(), UnusedUpdater(), UnusedWatcher(), query, UnusedMailSender(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal([id], query.GetArguments);
    }

    private static void AssertSummaryJson(string line, CalendarSummary expected)
    {
        using var json = JsonDocument.Parse(line);
        var root = json.RootElement;
        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(expected.Id, root.GetProperty("id").GetString());
        Assert.Equal(expected.Subject, root.GetProperty("subject").GetString());
        Assert.Equal(expected.StartAt, root.GetProperty("startAt").GetDateTimeOffset());
        Assert.Equal(expected.EndAt, root.GetProperty("endAt").GetDateTimeOffset());
        Assert.Equal(expected.IsAllDay, root.GetProperty("isAllDay").GetBoolean());
        Assert.Equal(expected.Location, root.GetProperty("location").GetString());
    }

    private static void AssertJsonEncoding(string line, string unicodeText, bool asciiOnly)
    {
        if (asciiOnly)
        {
            Assert.All(line, character => Assert.True(character <= '\u007F'));
            Assert.Contains("\\u", line, StringComparison.Ordinal);
        }
        else
        {
            Assert.Contains(unicodeText, line, StringComparison.Ordinal);
        }
    }

    private static List<string> ReadLines(string text)
    {
        using var reader = new StringReader(text);
        var lines = new List<string>();
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            lines.Add(line);
        }

        return lines;
    }

    private static IMailQuery UnusedMailQuery() =>
        new UnusedMailQueryStub();

    private static IMailReadStateUpdater UnusedUpdater() =>
        new UnusedUpdaterStub();

    private static IMailWatcher UnusedWatcher() =>
        new UnusedWatcherStub();

    private static IMailSender UnusedMailSender() =>
        new UnusedMailSenderStub();

    private sealed class RecordingCalendarQuery(
        Func<(DateTime From, DateTime To), IEnumerable<CalendarSummary>> listHandler,
        Func<string, CalendarDetail>? getHandler = null) : ICalendarQuery
    {
        public List<(DateTime From, DateTime To)> ListArguments { get; } = [];

        public List<string> GetArguments { get; } = [];

        public IEnumerable<CalendarSummary> List(DateTime from, DateTime to)
        {
            ListArguments.Add((from, to));
            return listHandler((from, to));
        }

        public CalendarDetail Get(string id)
        {
            GetArguments.Add(id);
            return (getHandler ?? throw new InvalidOperationException("Get must not be called"))(id);
        }
    }

    private sealed class UnusedMailQueryStub : IMailQuery
    {
        public IEnumerable<MailSummary> List(bool unreadOnly) =>
            throw new InvalidOperationException("Mail list must not be called");

        public MailDetail Get(string entryId) =>
            throw new InvalidOperationException("Mail show must not be called");
    }

    private sealed class UnusedUpdaterStub : IMailReadStateUpdater
    {
        public MailReadStateChange SetUnread(string entryId, bool unread) =>
            throw new InvalidOperationException("Mail state must not be called");
    }

    private sealed class UnusedWatcherStub : IMailWatcher
    {
        public void Watch(
            Action<MailSummary> onMail,
            Action onReady,
            CancellationToken cancellationToken) =>
            throw new InvalidOperationException("Mail watch must not be called");
    }

    private sealed class UnusedMailSenderStub : IMailSender
    {
        public void Send(string to, string subject, string body) =>
            throw new InvalidOperationException("Mail send must not be called");
    }
}
