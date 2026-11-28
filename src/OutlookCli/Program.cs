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
        var mail = new OutlookMailQuery();
        var calendar = new OutlookCalendarQuery();
        var services = new CliServices(
            MailQuery: mail,
            MailReadStateUpdater: new OutlookMailReadStateUpdater(),
            MailWatcher: new OutlookMailWatcher(),
            CalendarQuery: calendar,
            MailSender: new OutlookMailSender(),
            MailOperations: mail,
            CalendarOperations: calendar,
            ContactService: new OutlookContactService(),
            TaskService: new OutlookTaskService());
        return Run(args, services, Console.Out, Console.Error);
    }

    internal static int Run(
        string[] args,
        CliServices services,
        TextWriter output,
        TextWriter error)
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
        mail.Subcommands.Add(new MailListCommand(services.MailQuery, json, error).Create());
        mail.Subcommands.Add(new MailShowCommand(services.MailQuery, json, error).Create());
        mail.Subcommands.Add(new MailWatchCommand(services.MailWatcher, json, error).Create());
        var readState = new MailReadStateCommand(services.MailReadStateUpdater, json, error);
        mail.Subcommands.Add(readState.Create(unread: false));
        mail.Subcommands.Add(readState.Create(unread: true));
        mail.Subcommands.Add(new MailSendCommand(services.MailSender, json, error).Create());
        var mailFolders = new MailFolderCommand(services.MailOperations, json, error).Create();
        mailFolders.SetAction(help.Invoke);
        mail.Subcommands.Add(mailFolders);
        var mailAttachments = new MailAttachmentCommand(services.MailOperations, json, error).Create();
        mailAttachments.SetAction(help.Invoke);
        mail.Subcommands.Add(mailAttachments);
        mail.Subcommands.Add(new MailMoveCommand(services.MailOperations, json, error).Create());
        mail.Subcommands.Add(new MailReplyCommand(services.MailOperations, json, error).Create());
        root.Subcommands.Add(mail);

        var calendar = new Command("calendar", "Outlook の予定表を操作します。");
        calendar.SetAction(help.Invoke);
        calendar.Subcommands.Add(new CalendarListCommand(services.CalendarQuery, json, error).Create());
        calendar.Subcommands.Add(new CalendarShowCommand(services.CalendarQuery, json, error).Create());
        calendar.Subcommands.Add(new CalendarCreateCommand(services.CalendarOperations, json, error).Create());
        calendar.Subcommands.Add(new CalendarFreeBusyCommand(services.CalendarOperations, json, error).Create());
        root.Subcommands.Add(calendar);

        var contact = new ContactCommand(services.ContactService, json, error).Create();
        contact.SetAction(help.Invoke);
        root.Subcommands.Add(contact);

        var task = new TaskCommand(services.TaskService, json, error).Create();
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

    internal sealed record CliServices(
        IMailQuery MailQuery,
        IMailReadStateUpdater MailReadStateUpdater,
        IMailWatcher MailWatcher,
        ICalendarQuery CalendarQuery,
        IMailSender MailSender,
        IMailOperations MailOperations,
        ICalendarOperations CalendarOperations,
        IContactService ContactService,
        ITaskService TaskService);
}
