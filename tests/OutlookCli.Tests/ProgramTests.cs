using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Calendar;
using OutlookCli.Core.Mail;
using Xunit;

namespace OutlookCli.Tests;

public sealed class ProgramTests
{
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Command_group_without_subcommand_matches_explicit_help(int group)
    {
        string[] args = group switch
        {
            0 => [],
            1 => ["mail"],
            2 => ["calendar"],
            _ => throw new ArgumentOutOfRangeException(nameof(group)),
        };
        var query = new RecordingMailQuery(_ => throw new InvalidOperationException("List must not be called"));
        var output = new StringWriter();
        var explicitHelp = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(args, query, UnusedUpdater(), UnusedWatcher(), output, error);
        var helpExitCode = Run([.. args, "-h"], query, UnusedUpdater(), UnusedWatcher(), explicitHelp, error);

        Assert.Equal(0, exitCode);
        Assert.Equal(0, helpExitCode);
        Assert.NotEmpty(explicitHelp.ToString());
        Assert.Equal(explicitHelp.ToString(), output.ToString());
        Assert.Empty(error.ToString());
        Assert.False(query.WasCalled);
    }

    [Fact]
    public void Send_passes_the_message_to_outlook_and_reports_submission()
    {
        const string to = "recipient@example.com";
        const string subject = "日本語の件名";
        const string body = "1行目\n2行目";
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var sender = new RecordingMailSender((_, _, _) => { });
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(
            ["mail", "send", "--to", to, "--subject", subject, "--body", body],
            query, UnusedUpdater(), UnusedWatcher(), output, error, sender);

        Assert.Equal(0, exitCode);
        Assert.Equal((to, subject, body), Assert.Single(sender.Messages));
        Assert.Empty(error.ToString());
        using var json = JsonDocument.Parse(Assert.Single(ReadLines(output.ToString())));
        Assert.Equal(1, json.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.True(json.RootElement.GetProperty("submitted").GetBoolean());
    }

    [Theory]
    [InlineData("--to")]
    [InlineData("--subject")]
    [InlineData("--body")]
    public void Send_requires_all_message_options(string omittedOption)
    {
        var options = new Dictionary<string, string>
        {
            ["--to"] = "recipient@example.com",
            ["--subject"] = "Subject",
            ["--body"] = "Body",
        };
        options.Remove(omittedOption);
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var sender = new RecordingMailSender((_, _, _) => throw new InvalidOperationException("Send must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(
            ["mail", "send", .. options.SelectMany(option => new[] { option.Key, option.Value })],
            query, UnusedUpdater(), UnusedWatcher(), output, error, sender);

        Assert.Equal(1, exitCode);
        Assert.Empty(sender.Messages);
        Assert.Empty(output.ToString());
        Assert.Contains(omittedOption, error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Unread_option_passes_true_to_query()
    {
        var query = new RecordingMailQuery(_ =>
        [
            new MailSummary("件名", "送信者", new DateTimeOffset(2026, 11, 25, 9, 30, 0, TimeSpan.FromHours(9)), "entry-1")
        ]);
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list", "--unread"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Equal([true], query.UnreadOnlyArguments);
        Assert.Empty(error.ToString());
    }

    [Fact]
    public void Without_unread_option_passes_false_to_query()
    {
        var query = new RecordingMailQuery(_ =>
        [
            new MailSummary("件名", "送信者", new DateTimeOffset(2026, 11, 25, 9, 30, 0, TimeSpan.FromHours(9)), "entry-1")
        ]);
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Equal([false], query.UnreadOnlyArguments);
        Assert.Empty(error.ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Serializes_each_message_as_one_json_line_with_required_fields(bool asciiOnly)
    {
        var receivedAt = new DateTimeOffset(2026, 11, 25, 9, 30, 0, TimeSpan.FromHours(9));
        const string entryId = "entry-1";
        const string subject = "日本語の件名\n二行目";
        const string senderName = "山田 太郎";
        var query = new RecordingMailQuery(_ =>
        [
            new MailSummary(subject, senderName, receivedAt, entryId),
            new MailSummary("別の件名", senderName, receivedAt.AddMinutes(-5), "entry-2"),
        ]);
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly ? ["--ascii", "mail", "list"] : ["mail", "list"];
        var exitCode = Run(args, query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());

        var lines = ReadLines(output.ToString());
        Assert.Equal(2, lines.Count);
        AssertJsonEncoding(lines[0], senderName, asciiOnly);

        using var json = JsonDocument.Parse(lines[0]);
        var root = json.RootElement;
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(receivedAt, root.GetProperty("receivedAt").GetDateTimeOffset());
        Assert.Equal(senderName, root.GetProperty("senderName").GetString());
        Assert.Equal(subject, root.GetProperty("subject").GetString());
        Assert.Equal(entryId, root.GetProperty("entryId").GetString());

        using var secondJson = JsonDocument.Parse(lines[1]);
        Assert.Equal("別の件名", secondJson.RootElement.GetProperty("subject").GetString());
    }

    [Fact]
    public void Empty_result_returns_zero_and_writes_no_stdout()
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(output.ToString());
        Assert.Empty(error.ToString());
        Assert.Equal([false], query.UnreadOnlyArguments);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Watch_writes_ready_status_to_stderr_and_serializes_notifications_in_order(bool asciiOnly)
    {
        var firstReceivedAt = new DateTimeOffset(2026, 11, 25, 9, 30, 0, TimeSpan.FromHours(9));
        var first = new MailSummary("日本語の件名📬\n改行付き", "山田 太郎", firstReceivedAt, "entry-1");
        var secondReceivedAt = firstReceivedAt.AddMinutes(-5);
        var second = new MailSummary("second subject", "Second sender", secondReceivedAt, "entry-2");
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var watcher = new RecordingMailWatcher((onMail, onReady, cancellationToken) =>
        {
            Assert.True(cancellationToken.CanBeCanceled);
            onReady();
            onMail(first);
            onMail(second);
        });
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly ? ["mail", "watch", "--ascii"] : ["mail", "watch"];
        var exitCode = Run(args, query, UnusedUpdater(), watcher, output, error);

        Assert.Equal(0, exitCode);
        Assert.False(query.WasCalled);
        Assert.True(watcher.CancellationToken.CanBeCanceled);
        Assert.Contains("購読しました", error.ToString(), StringComparison.Ordinal);

        var lines = ReadLines(output.ToString());
        Assert.Equal(2, lines.Count);
        AssertJsonEncoding(lines[0], "日本語の件名📬", asciiOnly);

        using var firstJson = JsonDocument.Parse(lines[0]);
        var firstRoot = firstJson.RootElement;
        Assert.Equal(1, firstRoot.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(first.EntryId, firstRoot.GetProperty("entryId").GetString());
        Assert.Equal(first.ReceivedAt, firstRoot.GetProperty("receivedAt").GetDateTimeOffset());
        Assert.Equal(first.SenderName, firstRoot.GetProperty("senderName").GetString());
        Assert.Equal(first.Subject, firstRoot.GetProperty("subject").GetString());

        using var secondJson = JsonDocument.Parse(lines[1]);
        var secondRoot = secondJson.RootElement;
        Assert.Equal(1, secondRoot.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(second.EntryId, secondRoot.GetProperty("entryId").GetString());
        Assert.Equal(second.ReceivedAt, secondRoot.GetProperty("receivedAt").GetDateTimeOffset());
        Assert.Equal(second.SenderName, secondRoot.GetProperty("senderName").GetString());
        Assert.Equal(second.Subject, secondRoot.GetProperty("subject").GetString());
    }

    [Fact]
    public void Watch_without_notifications_returns_zero_and_writes_no_stdout()
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var watcher = new RecordingMailWatcher((_, onReady, cancellationToken) =>
        {
            Assert.True(cancellationToken.CanBeCanceled);
            onReady();
        });
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "watch"], query, UnusedUpdater(), watcher, output, error);

        Assert.Equal(0, exitCode);
        Assert.False(query.WasCalled);
        Assert.Empty(output.ToString());
        Assert.Contains("購読しました", error.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void Watch_com_exception_returns_one_and_reports_hresult()
    {
        const int hResult = unchecked((int)0x80004005);
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var watcher = new RecordingMailWatcher((_, _, _) =>
            throw new COMException("Outlook failed", hResult));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "watch"], query, UnusedUpdater(), watcher, output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.False(query.WasCalled);
    }

    [Fact]
    public void Com_exception_returns_one_and_reports_hresult()
    {
        const int hResult = unchecked((int)0x80004005);
        var query = new RecordingMailQuery(_ => throw new COMException("Outlook failed", hResult));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal([false], query.UnreadOnlyArguments);
    }

    [Fact]
    public void Com_exception_during_enumeration_preserves_lines_already_written()
    {
        var query = new RecordingMailQuery(_ => ReadThenFail());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list", "--unread"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        using var json = JsonDocument.Parse(Assert.Single(ReadLines(output.ToString())));
        Assert.Equal("取得済み", json.RootElement.GetProperty("subject").GetString());

        static IEnumerable<MailSummary> ReadThenFail()
        {
            yield return new MailSummary("取得済み", "差出人", DateTimeOffset.FromUnixTimeSeconds(0), "entry-1");
            throw new COMException("Outlook failed", unchecked((int)0x80004005));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Show_calls_get_once_and_serializes_one_json_object(bool asciiOnly)
    {
        const string entryId = "entry-42";
        const string subject = "Subject";
        const string senderName = "Sender";
        var receivedAt = new DateTimeOffset(2026, 11, 25, 9, 30, 0, TimeSpan.FromHours(9));
        const string to = "to@example.com";
        const string cc = "cc@example.com";
        const string bcc = "bcc@example.com";
        const string body = "本文\n引用符\"とバックスラッシュ\\u672C";
        var summary = new MailSummary(subject, senderName, receivedAt, entryId);
        var query = new RecordingMailQuery(
            _ => Array.Empty<MailSummary>(),
            _ => new MailDetail(summary, to, cc, bcc, body));
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly ? ["mail", "--ascii", "show", "--entry-id", entryId] : ["mail", "show", "--entry-id", entryId];
        var exitCode = Run(args, query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal([entryId], query.GetEntryIdArguments);

        var line = Assert.Single(ReadLines(output.ToString()));
        AssertJsonEncoding(line, "本文", asciiOnly);
        using var json = JsonDocument.Parse(line);
        var root = json.RootElement;
        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(entryId, root.GetProperty("entryId").GetString());
        Assert.Equal(subject, root.GetProperty("subject").GetString());
        Assert.Equal(senderName, root.GetProperty("senderName").GetString());
        Assert.Equal(receivedAt, root.GetProperty("receivedAt").GetDateTimeOffset());
        Assert.Equal(to, root.GetProperty("to").GetString());
        Assert.Equal(cc, root.GetProperty("cc").GetString());
        Assert.Equal(bcc, root.GetProperty("bcc").GetString());
        Assert.Equal(body, root.GetProperty("body").GetString());
    }

    [Fact]
    public void Show_without_entry_id_does_not_call_get_and_returns_one()
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "show"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.GetEntryIdArguments);
    }

    [Fact]
    public void Show_with_whitespace_entry_id_does_not_call_get_and_returns_one()
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "show", "--entry-id", "   "], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(query.GetEntryIdArguments);
    }

    [Fact]
    public void Show_com_exception_returns_one_and_reports_hresult()
    {
        const string entryId = "entry-42";
        var query = new RecordingMailQuery(
            _ => Array.Empty<MailSummary>(),
            _ => throw new COMException("Outlook failed", unchecked((int)0x80004005)));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "show", "--entry-id", entryId], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal([entryId], query.GetEntryIdArguments);
    }

    [Theory]
    [InlineData("mark-read", false, true, false)]
    [InlineData("mark-read", false, false, true)]
    [InlineData("mark-unread", true, true, false)]
    [InlineData("mark-unread", true, false, true)]
    public void Mark_state_calls_set_unread_once_and_serializes_result(
        string commandName,
        bool unread,
        bool changed,
        bool asciiOnly)
    {
        const string entryId = "entry-42";
        const string returnedEntryId = "保存済みID";
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var updater = new RecordingMailReadStateUpdater(
            (_, _) => new MailReadStateChange(returnedEntryId, changed));
        var output = new StringWriter();
        var error = new StringWriter();

        string[] args = asciiOnly ? ["mail", commandName, "--entry-id", entryId, "--ascii"] : ["mail", commandName, "--entry-id", entryId];
        var exitCode = Run(args, query, updater, UnusedWatcher(), output, error);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.False(query.WasCalled);
        Assert.Equal([entryId], updater.EntryIdArguments);
        Assert.Equal([unread], updater.UnreadArguments);

        var line = Assert.Single(ReadLines(output.ToString()));
        AssertJsonEncoding(line, returnedEntryId, asciiOnly);
        using var json = JsonDocument.Parse(line);
        var root = json.RootElement;
        Assert.Equal(JsonValueKind.Object, root.ValueKind);
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(returnedEntryId, root.GetProperty("entryId").GetString());
        Assert.Equal(unread, root.GetProperty("unread").GetBoolean());
        Assert.Equal(changed, root.GetProperty("changed").GetBoolean());
    }

    [Theory]
    [InlineData("mark-read")]
    [InlineData("mark-unread")]
    public void Mark_state_without_entry_id_does_not_call_set_unread_and_returns_one(string commandName)
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var updater = new RecordingMailReadStateUpdater(
            (_, _) => throw new InvalidOperationException("SetUnread must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(
            ["mail", commandName], query, updater, UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(updater.EntryIdArguments);
    }

    [Theory]
    [InlineData("mark-read")]
    [InlineData("mark-unread")]
    public void Mark_state_with_whitespace_entry_id_does_not_call_set_unread_and_returns_one(string commandName)
    {
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var updater = new RecordingMailReadStateUpdater(
            (_, _) => throw new InvalidOperationException("SetUnread must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(
            ["mail", commandName, "--entry-id", "   "], query, updater, UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Empty(updater.EntryIdArguments);
    }

    [Theory]
    [InlineData("mark-read", false)]
    [InlineData("mark-unread", true)]
    public void Mark_state_com_exception_returns_one_and_reports_hresult(string commandName, bool unread)
    {
        const string entryId = "entry-42";
        var query = new RecordingMailQuery(_ => Array.Empty<MailSummary>());
        var updater = new RecordingMailReadStateUpdater(
            (_, _) => throw new COMException("Outlook failed", unchecked((int)0x80004005)));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(
            ["mail", commandName, "--entry-id", entryId], query, updater, UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal([entryId], updater.EntryIdArguments);
        Assert.Equal([unread], updater.UnreadArguments);
    }

    [Fact]
    public void Parse_error_does_not_call_mail_query()
    {
        var query = new RecordingMailQuery(_ => throw new InvalidOperationException("query must not be called"));
        var output = new StringWriter();
        var error = new StringWriter();

        var exitCode = Run(["mail", "list", "--unknown"], query, UnusedUpdater(), UnusedWatcher(), output, error);

        Assert.Equal(1, exitCode);
        Assert.False(query.WasCalled);
        Assert.NotEmpty(error.ToString());
        Assert.Empty(output.ToString());
    }

    private static int Run(
        string[] args,
        IMailQuery query,
        IMailReadStateUpdater updater,
        IMailWatcher watcher,
        TextWriter output,
        TextWriter error,
        IMailSender? sender = null) =>
        OutlookCli.Program.Run(
            args, query, updater, watcher, new UnusedCalendarQuery(), sender ?? UnusedSender(), output, error);

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

    private static RecordingMailReadStateUpdater UnusedUpdater() =>
        new((_, _) => throw new InvalidOperationException("SetUnread must not be called"));

    private static RecordingMailWatcher UnusedWatcher() =>
        new((_, _, _) => throw new InvalidOperationException("Watch must not be called"));

    private static RecordingMailSender UnusedSender() =>
        new((_, _, _) => throw new InvalidOperationException("Send must not be called"));

    private sealed class UnusedCalendarQuery : ICalendarQuery
    {
        public IEnumerable<CalendarSummary> List(DateTime from, DateTime to) =>
            throw new NotSupportedException();

        public CalendarDetail Get(string id) => throw new NotSupportedException();
    }

    private sealed class RecordingMailQuery(
        Func<bool, IEnumerable<MailSummary>> listHandler,
        Func<string, MailDetail>? getHandler = null) : IMailQuery
    {
        public List<bool> UnreadOnlyArguments { get; } = [];

        public List<string> GetEntryIdArguments { get; } = [];

        public bool WasCalled => UnreadOnlyArguments.Count > 0 || GetEntryIdArguments.Count > 0;

        public IEnumerable<MailSummary> List(bool unreadOnly)
        {
            UnreadOnlyArguments.Add(unreadOnly);
            return listHandler(unreadOnly);
        }

        public MailDetail Get(string entryId)
        {
            GetEntryIdArguments.Add(entryId);
            return (getHandler ?? throw new InvalidOperationException("Get must not be called"))(entryId);
        }
    }

    private sealed class RecordingMailReadStateUpdater(
        Func<string, bool, MailReadStateChange> setUnreadHandler) : IMailReadStateUpdater
    {
        public List<string> EntryIdArguments { get; } = [];

        public List<bool> UnreadArguments { get; } = [];

        public MailReadStateChange SetUnread(string entryId, bool unread)
        {
            EntryIdArguments.Add(entryId);
            UnreadArguments.Add(unread);
            return setUnreadHandler(entryId, unread);
        }
    }

    private sealed class RecordingMailWatcher(
        Action<Action<MailSummary>, Action, CancellationToken> watchHandler) : IMailWatcher
    {
        public CancellationToken CancellationToken { get; private set; }

        public void Watch(
            Action<MailSummary> onMail,
            Action onReady,
            CancellationToken cancellationToken)
        {
            CancellationToken = cancellationToken;
            watchHandler(onMail, onReady, cancellationToken);
        }
    }

    private sealed class RecordingMailSender(Action<string, string, string> sendHandler) : IMailSender
    {
        public List<(string To, string Subject, string Body)> Messages { get; } = [];

        public void Send(string to, string subject, string body)
        {
            Messages.Add((to, subject, body));
            sendHandler(to, subject, body);
        }
    }
}
