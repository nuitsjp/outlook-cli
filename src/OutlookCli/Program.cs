using System.CommandLine;
using System.Text;
using OutlookCli.Commands;
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
            Console.Out, Console.Error);
    }

    internal static int Run(
        string[] args,
        IMailQuery query,
        IMailReadStateUpdater updater,
        IMailWatcher watcher,
        TextWriter output,
        TextWriter error)
    {
        var root = new RootCommand("クラシック Outlook のデータを取得します。");
        var mail = new Command("mail", "メールを操作します。");
        mail.Subcommands.Add(new MailListCommand(query, output, error).Create());
        mail.Subcommands.Add(new MailShowCommand(query, output, error).Create());
        mail.Subcommands.Add(new MailWatchCommand(watcher, output, error).Create());
        var readState = new MailReadStateCommand(updater, output, error);
        mail.Subcommands.Add(readState.Create(unread: false));
        mail.Subcommands.Add(readState.Create(unread: true));
        root.Subcommands.Add(mail);

        var parsed = root.Parse(args);
        if (parsed.Errors.Count > 0)
        {
            foreach (var parseError in parsed.Errors)
            {
                error.WriteLine(parseError.Message);
            }

            return 1;
        }

        return parsed.Invoke(new InvocationConfiguration { Output = output, Error = error });
    }
}
