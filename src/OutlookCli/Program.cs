using System.CommandLine;
using System.Text;
using OutlookCli.Commands;
using OutlookCli.Core.Calendar;
using OutlookCli.Core.Mail;
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
            new OutlookCalendarQuery(), Console.Out, Console.Error);
    }

    internal static int Run(
        string[] args,
        IMailQuery query,
        IMailReadStateUpdater updater,
        IMailWatcher watcher,
        ICalendarQuery calendarQuery,
        TextWriter output,
        TextWriter error)
    {
        var root = new RootCommand("クラシック Outlook のデータを取得します。");
        var ascii = new Option<bool>("--ascii")
        {
            Description = "JSON の非 ASCII 文字を Unicode エスケープで出力します。",
            Recursive = true,
        };
        root.Options.Add(ascii);

        var json = new JsonOutput(output);
        var mail = new Command("mail", "メールを操作します。");
        mail.Subcommands.Add(new MailListCommand(query, json, error).Create());
        mail.Subcommands.Add(new MailShowCommand(query, json, error).Create());
        mail.Subcommands.Add(new MailWatchCommand(watcher, json, error).Create());
        var readState = new MailReadStateCommand(updater, json, error);
        mail.Subcommands.Add(readState.Create(unread: false));
        mail.Subcommands.Add(readState.Create(unread: true));
        root.Subcommands.Add(mail);

        var calendar = new Command("calendar", "既定の予定表から予定を取得します。");
        calendar.Subcommands.Add(new CalendarListCommand(calendarQuery, json, error).Create());
        calendar.Subcommands.Add(new CalendarShowCommand(calendarQuery, json, error).Create());
        root.Subcommands.Add(calendar);

        root.Subcommands.Add(new UpdateCommand(json, error).Create());

        var parsed = root.Parse(args.Length == 0 ? ["-h"] : args);
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
