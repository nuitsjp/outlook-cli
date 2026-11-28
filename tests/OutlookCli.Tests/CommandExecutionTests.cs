using System.CommandLine;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OutlookCli.Commands;
using OutlookCli.Core.Mail;
using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class CommandExecutionTests
{
    [Fact]
    public void Reports_operation_hresult_and_each_cleanup_failure()
    {
        var operation = new COMException("reading item", unchecked((int)0x80004005));
        var release = new COMException("releasing folder", unchecked((int)0x80070005));
        var shutdown = new InvalidOperationException("waiting for Outlook");
        using var error = new StringWriter();

        var exitCode = CommandExecution.Run(error, "一覧の取得", () =>
            throw new OutlookComCleanupException(operation, [release, shutdown], false));

        Assert.Equal(1, exitCode);
        var messages = error.ToString().Split([Environment.NewLine], StringSplitOptions.RemoveEmptyEntries);
        Assert.Collection(messages,
            line =>
            {
                Assert.Contains("一覧の取得", line);
                Assert.Contains("0x80004005", line);
                Assert.Contains("reading item", line);
            },
            line =>
            {
                Assert.Contains("0x80070005", line);
                Assert.Contains("releasing folder", line);
            },
            line => Assert.Contains("waiting for Outlook", line));
    }

    [Fact]
    public void List_preserves_output_error_and_cleanup_diagnostic_after_a_partial_result()
    {
        static IEnumerable<int> ReadRows(ComCleanup cleanup)
        {
            try
            {
                yield return 1;
                yield return 2;
            }
            finally
            {
                cleanup.Attempt(() => throw new InvalidOperationException("closing folder"));
            }
        }

        using var output = new StringWriter();
        using var error = new StringWriter();
        var exitCode = CommandExecution.RunList(error, "一覧の取得", () => ComCleanup.Enumerate(ReadRows), row =>
        {
            if (row == 2)
            {
                throw new IOException("writing row");
            }
            output.WriteLine(row);
        });

        Assert.Equal(1, exitCode);
        Assert.Equal("1" + Environment.NewLine, output.ToString());
        Assert.Collection(error.ToString().Split([Environment.NewLine], StringSplitOptions.RemoveEmptyEntries),
            line => Assert.Contains("writing row", line),
            line => Assert.Contains("closing folder", line));
    }

    [Fact]
    public void Watch_flushes_each_notification_while_the_subscription_is_active()
    {
        using var stream = new MemoryStream();
        using var output = new StreamWriter(stream, new UTF8Encoding(false));
        using var error = new StringWriter();
        var watcher = new WatchingMailbox(stream);
        var root = new RootCommand();
        root.Subcommands.Add(new MailWatchCommand(watcher, new JsonOutput(output), error).Create());

        var exitCode = root.Parse(["watch"]).Invoke(new InvocationConfiguration { Output = output, Error = error });

        Assert.Equal(0, exitCode);
        Assert.Equal(["first", "second"], watcher.VisibleIds);
    }

    [Fact]
    public void Submitted_message_still_reports_success_when_releasing_outlook_fails()
    {
        using var output = new StringWriter();
        using var error = new StringWriter();
        var sender = new SubmittedMailSender();
        var root = new RootCommand();
        root.Subcommands.Add(new MailSendCommand(sender, new JsonOutput(output), error).Create());

        var exitCode = root.Parse(["send", "--to", "test@example.com", "--subject", "subject", "--body", "body"])
            .Invoke(new InvocationConfiguration { Output = output, Error = error });

        Assert.Equal(1, sender.Calls);
        Assert.Equal(0, exitCode);
        using var row = JsonDocument.Parse(output.ToString());
        Assert.True(row.RootElement.GetProperty("submitted").GetBoolean());
        Assert.Contains("releasing recipients", error.ToString());
    }

    private sealed class SubmittedMailSender : IMailSender
    {
        public int Calls { get; private set; }

        public void Send(string to, string subject, string body)
        {
            Calls++;
            throw new OutlookComCleanupException(null, [new InvalidOperationException("releasing recipients")], true);
        }
    }

    private sealed class WatchingMailbox(MemoryStream output) : IMailWatcher
    {
        public List<string> VisibleIds { get; } = [];

        public void Watch(Action<MailSummary> onMail, Action onReady, CancellationToken cancellationToken)
        {
            onReady();
            foreach (var id in new[] { "first", "second" })
            {
                var previousLength = (int)output.Length;
                onMail(new MailSummary("日本語", "", DateTimeOffset.FromUnixTimeSeconds(0), id));
                var bytes = output.ToArray();
                Assert.True(bytes.Length > previousLength);
                var notificationBytes = new byte[bytes.Length - previousLength];
                Array.Copy(bytes, previousLength, notificationBytes, 0, notificationBytes.Length);
                using var row = JsonDocument.Parse(notificationBytes);
                VisibleIds.Add(row.RootElement.GetProperty("entryId").GetString()!);
            }
        }
    }
}
