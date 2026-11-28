using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookCli.Core.Calendar;
using OutlookCli.Core.Contacts;
using OutlookCli.Core.Mail;
using OutlookCli.Core.Tasks;
using Xunit;

namespace OutlookCli.Tests;

public sealed class ExtendedCommandTests
{
    [Fact]
    public void Mail_folder_list_writes_folder_json()
    {
        var operations = new RecordingMailOperations
        {
            Folders = [new MailFolderSummary("folder-1", "受信トレイ", "root", "\\受信トレイ", "store-1")],
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(["mail", "folder", "list"], output, error, mailOperations: operations);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(1, operations.ListFoldersCalls);
        var json = ParseSingle(output.ToString());
        Assert.Equal(1, json.GetProperty("schemaVersion").GetInt32());
        Assert.Equal("folder-1", json.GetProperty("folderId").GetString());
        Assert.Equal("受信トレイ", json.GetProperty("displayName").GetString());
        Assert.Equal("root", json.GetProperty("parentFolderId").GetString());
        Assert.Equal("store-1", json.GetProperty("storeId").GetString());
    }

    [Fact]
    public void Mail_attachment_list_writes_attachment_json()
    {
        var operations = new RecordingMailOperations
        {
            Attachments = [new MailAttachmentSummary("1", "report.pdf", 42, 1)],
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["mail", "attachment", "list", "--entry-id", "mail-1"],
            output,
            error,
            mailOperations: operations);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(["mail-1"], operations.AttachmentListEntryIds);
        var json = ParseSingle(output.ToString());
        Assert.Equal("1", json.GetProperty("attachmentId").GetString());
        Assert.Equal("report.pdf", json.GetProperty("fileName").GetString());
        Assert.Equal(42L, json.GetProperty("size").GetInt64());
    }

    [Fact]
    public void Mail_attachment_save_passes_full_path_and_writes_result()
    {
        var path = Path.Combine("fixtures", $"outlook-cli-{Guid.NewGuid():N}", "report.pdf");
        var fullPath = Path.GetFullPath(path);
        var operations = new RecordingMailOperations
        {
            SavedAttachment = new MailAttachmentSaveResult("mail-1", "2", fullPath),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["mail", "attachment", "save", "--entry-id", "mail-1", "--attachment-id", "2", "--path", path],
            output,
            error,
            mailOperations: operations);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(operations.SavedAttachmentArguments.HasValue);
        Assert.Equal(("mail-1", "2", fullPath), operations.SavedAttachmentArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("mail-1", json.GetProperty("entryId").GetString());
        Assert.Equal("2", json.GetProperty("attachmentId").GetString());
        Assert.Equal(fullPath, json.GetProperty("path").GetString());
        Assert.True(json.GetProperty("saved").GetBoolean());
    }

    [Fact]
    public void Mail_attachment_save_rejects_existing_file_without_calling_service()
    {
        var path = Path.Combine(Path.GetTempPath(), $"outlook-cli-existing-{Guid.NewGuid():N}.txt");
        const string originalContent = "existing attachment content";
        File.WriteAllText(path, originalContent);
        var operations = new RecordingMailOperations();
        using var output = new StringWriter();
        using var error = new StringWriter();

        try
        {
            var exitCode = Run(
                ["mail", "attachment", "save", "--entry-id", "mail-1", "--attachment-id", "2", "--path", path],
                output,
                error,
                mailOperations: operations);

            Assert.Equal(1, exitCode);
            Assert.Empty(output.ToString());
            Assert.Contains(path, error.ToString(), StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, operations.TotalCalls);
            Assert.Null(operations.SavedAttachmentArguments);
            Assert.Equal(originalContent, File.ReadAllText(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Mail_attachment_save_rejects_path_too_long_without_calling_service()
    {
        var path = new string('a', 40_000);
        var operations = new RecordingMailOperations();
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["mail", "attachment", "save", "--entry-id", "mail-1", "--attachment-id", "2", "--path", path],
            output,
            error,
            mailOperations: operations);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Equal(0, operations.TotalCalls);
        Assert.Null(operations.SavedAttachmentArguments);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Mail_move_passes_ids_and_writes_changed_result(bool changed)
    {
        var operations = new RecordingMailOperations
        {
            MoveResult = new MailMoveResult("mail-2", "folder-2", "store-1", changed),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["mail", "move", "--entry-id", "mail-1", "--to-folder", "folder-2"],
            output,
            error,
            mailOperations: operations);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(operations.MoveArguments.HasValue);
        Assert.Equal(("mail-1", "folder-2"), operations.MoveArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("mail-2", json.GetProperty("entryId").GetString());
        Assert.Equal("folder-2", json.GetProperty("folderId").GetString());
        Assert.Equal(changed, json.GetProperty("changed").GetBoolean());
    }

    [Fact]
    public void Mail_reply_passes_ids_and_writes_draft_result()
    {
        var operations = new RecordingMailOperations
        {
            ReplyResult = new MailReplyResult("mail-1", "draft-1", "Re: Subject"),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["mail", "reply", "--entry-id", "mail-1", "--body", "確認しました"],
            output,
            error,
            mailOperations: operations);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(operations.ReplyArguments.HasValue);
        Assert.Equal(("mail-1", "確認しました"), operations.ReplyArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("draft-1", json.GetProperty("draftEntryId").GetString());
        Assert.Equal("Re: Subject", json.GetProperty("subject").GetString());
        Assert.True(json.GetProperty("saved").GetBoolean());
    }

    [Theory]
    [InlineData("folder list")]
    [InlineData("attachment list")]
    [InlineData("attachment save")]
    [InlineData("move")]
    [InlineData("reply")]
    public void Mail_operation_service_failure_returns_one(string command)
    {
        var operations = new RecordingMailOperations
        {
            Failure = new COMException("Outlook failed", unchecked((int)0x80004005)),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(MailOperationArguments(command), output, error, mailOperations: operations);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal(1, operations.TotalCalls);
    }

    [Fact]
    public void Contact_list_passes_filters_and_writes_summary()
    {
        var service = new RecordingContactService
        {
            Contacts = [new Contact("contact-1", "山田 太郎", "taro@example.com", "Example")],
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["contact", "list", "--name", "山田", "--company", "Example"],
            output,
            error,
            contactService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(service.ListArguments.HasValue);
        Assert.Equal(((string?)"山田", (string?)"Example"), service.ListArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("contact-1", json.GetProperty("entryId").GetString());
        Assert.Equal("taro@example.com", json.GetProperty("email").GetString());
    }

    [Fact]
    public void Contact_show_writes_detailed_contact()
    {
        var service = new RecordingContactService
        {
            Contact = new Contact("contact-1", "山田 太郎", "taro@example.com", "Example", "second@example.com", "", "03-0000-0000"),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(["contact", "show", "--entry-id", "contact-1"], output, error, contactService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(["contact-1"], service.GetEntryIds);
        var json = ParseSingle(output.ToString());
        Assert.Equal("second@example.com", json.GetProperty("email2").GetString());
        Assert.Equal("03-0000-0000", json.GetProperty("businessPhone").GetString());
    }

    [Fact]
    public void Contact_add_passes_values_and_writes_contact()
    {
        var service = new RecordingContactService
        {
            Contact = new Contact("contact-2", "山田 太郎", "taro@example.com", "Example"),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["contact", "add", "--name", "山田 太郎", "--email", "taro@example.com", "--company", "Example"],
            output,
            error,
            contactService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(service.AddArguments.HasValue);
        Assert.Equal(("山田 太郎", "taro@example.com", (string?)"Example"), service.AddArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("contact-2", json.GetProperty("entryId").GetString());
        Assert.Equal("Example", json.GetProperty("company").GetString());
    }

    [Fact]
    public void Contact_lookup_writes_address_book_result_without_entry_id()
    {
        var service = new RecordingContactService
        {
            Contact = new Contact("address-entry-1", "山田 太郎", "taro@example.com", "Example"),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(["contact", "lookup", "--name", "山田 太郎"], output, error, contactService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(["山田 太郎"], service.LookupNames);
        var json = ParseSingle(output.ToString());
        Assert.False(json.TryGetProperty("entryId", out _));
        Assert.Equal("taro@example.com", json.GetProperty("email").GetString());
    }

    [Theory]
    [InlineData("list")]
    [InlineData("show")]
    [InlineData("add")]
    [InlineData("lookup")]
    public void Contact_service_failure_returns_one(string command)
    {
        var service = new RecordingContactService
        {
            Failure = new COMException("Outlook failed", unchecked((int)0x80004005)),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(ContactArguments(command), output, error, contactService: service);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal(1, service.TotalCalls);
    }

    [Fact]
    public void Task_list_writes_summary()
    {
        var service = new RecordingTaskService
        {
            Tasks = [new TaskSummary("task-1", "確認", new DateTime(2026, 12, 1))],
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(["task", "list"], output, error, taskService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(1, service.ListCalls);
        var json = ParseSingle(output.ToString());
        Assert.Equal("task-1", json.GetProperty("entryId").GetString());
        Assert.Equal("確認", json.GetProperty("subject").GetString());
    }

    [Fact]
    public void Task_add_passes_subject_and_due()
    {
        var service = new RecordingTaskService
        {
            Task = new TaskSummary("task-2", "確認", new DateTime(2026, 12, 1)),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            ["task", "add", "--subject", "確認", "--due", "2026-12-01"],
            output,
            error,
            taskService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.True(service.AddArguments.HasValue);
        Assert.Equal(("確認", (DateTime?)new DateTime(2026, 12, 1)), service.AddArguments.Value);
        var json = ParseSingle(output.ToString());
        Assert.Equal("task-2", json.GetProperty("entryId").GetString());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Task_complete_passes_entry_id_and_writes_changed_result(bool changed)
    {
        var service = new RecordingTaskService
        {
            Change = new TaskChange("task-1", changed),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(["task", "complete", "--entry-id", "task-1"], output, error, taskService: service);

        Assert.Equal(0, exitCode);
        Assert.Empty(error.ToString());
        Assert.Equal(["task-1"], service.CompleteEntryIds);
        var json = ParseSingle(output.ToString());
        Assert.Equal("task-1", json.GetProperty("entryId").GetString());
        Assert.Equal(changed, json.GetProperty("changed").GetBoolean());
    }

    [Theory]
    [InlineData("list")]
    [InlineData("add")]
    [InlineData("complete")]
    public void Task_service_failure_returns_one(string command)
    {
        var service = new RecordingTaskService
        {
            Failure = new COMException("Outlook failed", unchecked((int)0x80004005)),
        };
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(TaskArguments(command), output, error, taskService: service);

        Assert.Equal(1, exitCode);
        Assert.Contains("80004005", error.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Empty(output.ToString());
        Assert.Equal(1, service.TotalCalls);
    }

    [Theory]
    [InlineData("mail attachment list")]
    [InlineData("mail attachment save")]
    [InlineData("mail move")]
    [InlineData("mail reply")]
    [InlineData("contact show")]
    [InlineData("contact add")]
    [InlineData("contact lookup")]
    [InlineData("task add")]
    [InlineData("task complete")]
    public void Required_options_reject_before_service_call(string command)
    {
        var mailOperations = new RecordingMailOperations();
        var contactService = new RecordingContactService();
        var taskService = new RecordingTaskService();
        using var output = new StringWriter();
        using var error = new StringWriter();

        var exitCode = Run(
            command.Split([' '], StringSplitOptions.RemoveEmptyEntries),
            output,
            error,
            mailOperations,
            contactService: contactService,
            taskService: taskService);

        Assert.Equal(1, exitCode);
        Assert.Empty(output.ToString());
        Assert.NotEmpty(error.ToString());
        Assert.Equal(0, mailOperations.TotalCalls);
        Assert.Equal(0, contactService.TotalCalls);
        Assert.Equal(0, taskService.TotalCalls);
    }

    private static int Run(
        string[] args,
        TextWriter output,
        TextWriter error,
        IMailOperations? mailOperations = null,
        ICalendarOperations? calendarOperations = null,
        IContactService? contactService = null,
        ITaskService? taskService = null) =>
        CliTestSupport.Run(
            args,
            new UnusedMailQuery(),
            new UnusedUpdater(),
            new UnusedWatcher(),
            new UnusedCalendarQuery(),
            new UnusedSender(),
            output,
            error,
            mailOperations,
            calendarOperations,
            contactService,
            taskService);

    private static string[] MailOperationArguments(string command) => command switch
    {
        "folder list" => ["mail", "folder", "list"],
        "attachment list" => ["mail", "attachment", "list", "--entry-id", "mail-1"],
        "attachment save" => ["mail", "attachment", "save", "--entry-id", "mail-1", "--attachment-id", "1", "--path", SavePath()],
        "move" => ["mail", "move", "--entry-id", "mail-1", "--to-folder", "folder-1"],
        "reply" => ["mail", "reply", "--entry-id", "mail-1", "--body", "本文"],
        _ => throw new ArgumentOutOfRangeException(nameof(command), command, null),
    };

    private static string[] ContactArguments(string command) => command switch
    {
        "list" => ["contact", "list"],
        "show" => ["contact", "show", "--entry-id", "contact-1"],
        "add" => ["contact", "add", "--name", "山田", "--email", "taro@example.com"],
        "lookup" => ["contact", "lookup", "--name", "山田"],
        _ => throw new ArgumentOutOfRangeException(nameof(command), command, null),
    };

    private static string[] TaskArguments(string command) => command switch
    {
        "list" => ["task", "list"],
        "add" => ["task", "add", "--subject", "確認"],
        "complete" => ["task", "complete", "--entry-id", "task-1"],
        _ => throw new ArgumentOutOfRangeException(nameof(command), command, null),
    };

    private static string SavePath() => Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "report.pdf");

    private static JsonElement ParseSingle(string text)
    {
        using var document = JsonDocument.Parse(text);
        return document.RootElement.Clone();
    }

    private sealed class RecordingMailOperations : IMailOperations
    {
        public IEnumerable<MailFolderSummary> Folders { get; init; } = [];
        public IEnumerable<MailAttachmentSummary> Attachments { get; init; } = [];
        public MailAttachmentSaveResult? SavedAttachment { get; init; }
        public MailMoveResult? MoveResult { get; init; }
        public MailReplyResult? ReplyResult { get; init; }
        public Exception? Failure { get; init; }
        public int ListFoldersCalls { get; private set; }
        public List<string> AttachmentListEntryIds { get; } = [];
        public (string EntryId, string AttachmentId, string Path)? SavedAttachmentArguments { get; private set; }
        public (string EntryId, string FolderId)? MoveArguments { get; private set; }
        public (string EntryId, string Body)? ReplyArguments { get; private set; }
        public int TotalCalls { get; private set; }

        public IEnumerable<MailFolderSummary> ListFolders()
        {
            ListFoldersCalls++;
            TotalCalls++;
            ThrowIfFailed();
            return Folders;
        }

        public IEnumerable<MailAttachmentSummary> ListAttachments(string entryId)
        {
            TotalCalls++;
            AttachmentListEntryIds.Add(entryId);
            ThrowIfFailed();
            return Attachments;
        }

        public MailAttachmentSaveResult SaveAttachment(string entryId, string attachmentId, string path)
        {
            TotalCalls++;
            SavedAttachmentArguments = (entryId, attachmentId, path);
            ThrowIfFailed();
            return SavedAttachment ?? throw new InvalidOperationException("SaveAttachment result was not configured");
        }

        public MailMoveResult Move(string entryId, string folderId)
        {
            TotalCalls++;
            MoveArguments = (entryId, folderId);
            ThrowIfFailed();
            return MoveResult ?? throw new InvalidOperationException("Move result was not configured");
        }

        public MailReplyResult Reply(string entryId, string body)
        {
            TotalCalls++;
            ReplyArguments = (entryId, body);
            ThrowIfFailed();
            return ReplyResult ?? throw new InvalidOperationException("Reply result was not configured");
        }

        private void ThrowIfFailed()
        {
            if (Failure is not null)
            {
                throw Failure;
            }
        }
    }

    private sealed class RecordingContactService : IContactService
    {
        public IEnumerable<Contact> Contacts { get; init; } = [];
        public Contact? Contact { get; init; }
        public Exception? Failure { get; init; }
        public (string? Name, string? Company)? ListArguments { get; private set; }
        public List<string> GetEntryIds { get; } = [];
        public (string Name, string Email, string? Company)? AddArguments { get; private set; }
        public List<string> LookupNames { get; } = [];
        public int TotalCalls { get; private set; }

        public IEnumerable<Contact> List(string? name, string? company)
        {
            TotalCalls++;
            ListArguments = (name, company);
            ThrowIfFailed();
            return Contacts;
        }

        public Contact Get(string entryId)
        {
            TotalCalls++;
            GetEntryIds.Add(entryId);
            ThrowIfFailed();
            return Contact ?? throw new InvalidOperationException("Contact result was not configured");
        }

        public Contact Add(string name, string email, string? company)
        {
            TotalCalls++;
            AddArguments = (name, email, company);
            ThrowIfFailed();
            return Contact ?? throw new InvalidOperationException("Contact result was not configured");
        }

        public Contact Lookup(string name)
        {
            TotalCalls++;
            LookupNames.Add(name);
            ThrowIfFailed();
            return Contact ?? throw new InvalidOperationException("Contact result was not configured");
        }

        private void ThrowIfFailed()
        {
            if (Failure is not null)
            {
                throw Failure;
            }
        }
    }

    private sealed class RecordingTaskService : ITaskService
    {
        public IEnumerable<TaskSummary> Tasks { get; init; } = [];
        public TaskSummary? Task { get; init; }
        public TaskChange? Change { get; init; }
        public Exception? Failure { get; init; }
        public int ListCalls { get; private set; }
        public (string Subject, DateTime? Due)? AddArguments { get; private set; }
        public List<string> CompleteEntryIds { get; } = [];
        public int TotalCalls { get; private set; }

        public IEnumerable<TaskSummary> List()
        {
            ListCalls++;
            TotalCalls++;
            ThrowIfFailed();
            return Tasks;
        }

        public TaskSummary Add(string subject, DateTime? due)
        {
            TotalCalls++;
            AddArguments = (subject, due);
            ThrowIfFailed();
            return Task ?? throw new InvalidOperationException("Task result was not configured");
        }

        public TaskChange Complete(string entryId)
        {
            TotalCalls++;
            CompleteEntryIds.Add(entryId);
            ThrowIfFailed();
            return Change ?? throw new InvalidOperationException("Task change was not configured");
        }

        private void ThrowIfFailed()
        {
            if (Failure is not null)
            {
                throw Failure;
            }
        }
    }

    private sealed class UnusedMailQuery : IMailQuery
    {
        public IEnumerable<MailSummary> List(MailListFilter filter) =>
            throw new InvalidOperationException("Mail query must not be called");

        public MailDetail Get(string entryId) =>
            throw new InvalidOperationException("Mail query must not be called");
    }

    private sealed class UnusedUpdater : IMailReadStateUpdater
    {
        public MailReadStateChange SetUnread(string entryId, bool unread) =>
            throw new InvalidOperationException("Read state updater must not be called");
    }

    private sealed class UnusedWatcher : IMailWatcher
    {
        public void Watch(Action<MailSummary> onMail, Action onReady, CancellationToken cancellationToken) =>
            throw new InvalidOperationException("Mail watcher must not be called");
    }

    private sealed class UnusedCalendarQuery : ICalendarQuery
    {
        public IEnumerable<CalendarSummary> List(DateTime from, DateTime to, string? owner = null) =>
            throw new InvalidOperationException("Calendar query must not be called");

        public CalendarDetail Get(string id, string? owner = null) =>
            throw new InvalidOperationException("Calendar query must not be called");
    }

    private sealed class UnusedSender : IMailSender
    {
        public void Send(string to, string subject, string body) =>
            throw new InvalidOperationException("Mail sender must not be called");
    }
}
