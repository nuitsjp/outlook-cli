using OutlookCli.Core.Calendar;
using OutlookCli.Core.Contacts;
using OutlookCli.Core.Mail;
using OutlookCli.Core.Tasks;

namespace OutlookCli.Tests;

internal static class CliTestSupport
{
    public static int Run(
        string[] args,
        IMailQuery query,
        IMailReadStateUpdater updater,
        IMailWatcher watcher,
        ICalendarQuery calendarQuery,
        IMailSender sender,
        TextWriter output,
        TextWriter error,
        IMailOperations? mailOperations = null,
        ICalendarOperations? calendarOperations = null,
        IContactService? contactService = null,
        ITaskService? taskService = null) =>
        OutlookCli.Program.Run(
            args,
            new OutlookCli.Program.CliServices(
                MailQuery: query,
                MailReadStateUpdater: updater,
                MailWatcher: watcher,
                CalendarQuery: calendarQuery,
                MailSender: sender,
                MailOperations: mailOperations ?? new UnusedMailOperations(),
                CalendarOperations: calendarOperations ?? new UnusedCalendarOperations(),
                ContactService: contactService ?? new UnusedContactService(),
                TaskService: taskService ?? new UnusedTaskService()),
            output,
            error);

    private sealed class UnusedMailOperations : IMailOperations
    {
        public IEnumerable<MailFolderSummary> ListFolders() =>
            throw new InvalidOperationException("Mail operations must not be called");

        public IEnumerable<MailAttachmentSummary> ListAttachments(string entryId) =>
            throw new InvalidOperationException("Mail operations must not be called");

        public MailAttachmentSaveResult SaveAttachment(string entryId, string attachmentId, string path) =>
            throw new InvalidOperationException("Mail operations must not be called");

        public MailMoveResult Move(string entryId, string folderId) =>
            throw new InvalidOperationException("Mail operations must not be called");

        public MailReplyResult Reply(string entryId, string body) =>
            throw new InvalidOperationException("Mail operations must not be called");
    }

    private sealed class UnusedCalendarOperations : ICalendarOperations
    {
        public CalendarSummary Create(string subject, DateTimeOffset start, DateTimeOffset end) =>
            throw new InvalidOperationException("Calendar operations must not be called");

        public IEnumerable<CalendarAvailability> GetAvailability(
            IReadOnlyList<string> attendees,
            DateTime from,
            DateTime to) =>
            throw new InvalidOperationException("Calendar operations must not be called");
    }

    private sealed class UnusedContactService : IContactService
    {
        public IEnumerable<Contact> List(string? name, string? company) =>
            throw new InvalidOperationException("Contact service must not be called");

        public Contact Get(string entryId) =>
            throw new InvalidOperationException("Contact service must not be called");

        public Contact Add(string name, string email, string? company) =>
            throw new InvalidOperationException("Contact service must not be called");

        public Contact Lookup(string name) =>
            throw new InvalidOperationException("Contact service must not be called");
    }

    private sealed class UnusedTaskService : ITaskService
    {
        public IEnumerable<TaskSummary> List() =>
            throw new InvalidOperationException("Task service must not be called");

        public TaskSummary Add(string subject, DateTime? due) =>
            throw new InvalidOperationException("Task service must not be called");

        public TaskChange Complete(string entryId) =>
            throw new InvalidOperationException("Task service must not be called");
    }
}
