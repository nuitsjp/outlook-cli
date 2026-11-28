using System.CommandLine;
using System.CommandLine.Help;
using System.Text;
using OutlookCli.Commands;
using OutlookCli.Core.Calendar;
using OutlookCli.Core.Contacts;
using OutlookCli.Core.Mail;
using OutlookCli.Core.Tasks;
using OutlookCli.OutlookCom;

namespace OutlookCli;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return Run(
            args, new OutlookMailQuery(), new OutlookMailReadStateUpdater(), new OutlookMailWatcher(),
            new OutlookCalendarQuery(), new OutlookMailSender(), Console.Out, Console.Error);
    }

    internal static int Run(
        string[] args,
        IMailQuery query,
        IMailReadStateUpdater updater,
        IMailWatcher watcher,
        ICalendarQuery calendarQuery,
        IMailSender sender,
        TextWriter output,
        TextWriter error,
        ICalendarOperations? calendarOperations = null,
        IContactService? contactService = null,
        ITaskService? taskService = null,
        IMailOperations? mailOperations = null)
    {
        var root = new RootCommand("クラシック Outlook を操作します。");
        var ascii = new Option<bool>("--ascii")
        {
            Description = "JSON の非 ASCII 文字を Unicode エスケープで出力します。",
            Recursive = true,
        };
        root.Options.Add(ascii);

        var json = new JsonOutput(output);
        var mail = new Command("mail", "メールを操作します。");
        var help = (HelpAction)root.Options.OfType<HelpOption>().Single().Action!;
        root.SetAction(help.Invoke);
        mail.SetAction(help.Invoke);
        mail.Subcommands.Add(new MailListCommand(query, json, error).Create());
        mail.Subcommands.Add(new MailShowCommand(query, json, error).Create());
        mail.Subcommands.Add(new MailWatchCommand(watcher, json, error).Create());
        var readState = new MailReadStateCommand(updater, json, error);
        mail.Subcommands.Add(readState.Create(unread: false));
        mail.Subcommands.Add(readState.Create(unread: true));
        mail.Subcommands.Add(new MailSendCommand(sender, json, error).Create());
        var mailActions = mailOperations ?? new OutlookMailQuery();
        var mailFolders = new MailFolderCommand(mailActions, json, error).Create();
        mailFolders.SetAction(help.Invoke);
        mail.Subcommands.Add(mailFolders);
        var mailAttachments = new MailAttachmentCommand(mailActions, json, error).Create();
        mailAttachments.SetAction(help.Invoke);
        mail.Subcommands.Add(mailAttachments);
        mail.Subcommands.Add(new MailMoveCommand(mailActions, json, error).Create());
        mail.Subcommands.Add(new MailReplyCommand(mailActions, json, error).Create());
        root.Subcommands.Add(mail);

        var calendar = new Command("calendar", "Outlook の予定表を操作します。");
        calendar.SetAction(help.Invoke);
        calendar.Subcommands.Add(new CalendarListCommand(calendarQuery, json, error).Create());
        calendar.Subcommands.Add(new CalendarShowCommand(calendarQuery, json, error).Create());
        var operations = calendarOperations ?? new OutlookCalendarQuery();
        calendar.Subcommands.Add(new CalendarCreateCommand(operations, json, error).Create());
        calendar.Subcommands.Add(new CalendarFreeBusyCommand(operations, json, error).Create());
        root.Subcommands.Add(calendar);

        var contact = new ContactCommand(contactService ?? new OutlookContactService(), json, error).Create();
        contact.SetAction(help.Invoke);
        root.Subcommands.Add(contact);

        var task = new TaskCommand(taskService ?? new OutlookTaskService(), json, error).Create();
        task.SetAction(help.Invoke);
        root.Subcommands.Add(task);

        root.Subcommands.Add(new UpdateCommand(json, error).Create());

        var parsed = root.Parse(args);
        if (parsed.Errors.Count > 0)
        {
            foreach (var parseError in parsed.Errors)
            {
                error.WriteLine(parseError.Message);
            }

            return 1;
        }

        json.AsciiOnly = parsed.GetValue(ascii);
        return parsed.Invoke(new InvocationConfiguration { Output = output, Error = error });
    }
}
