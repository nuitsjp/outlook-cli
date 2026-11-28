using System.CommandLine;
using OutlookCli.Core.Contacts;

namespace OutlookCli.Commands;

internal sealed class ContactCommand(IContactService service, JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("contact", "連絡先を操作します。");
        command.Subcommands.Add(CreateList());
        command.Subcommands.Add(CreateShow());
        command.Subcommands.Add(CreateAdd());
        command.Subcommands.Add(CreateLookup());
        return command;
    }

    private Command CreateList()
    {
        var name = new Option<string>("--name") { Description = "連絡先名の一部で絞り込みます。" };
        var company = new Option<string>("--company") { Description = "会社名の一部で絞り込みます。" };
        var command = new Command("list", "既定の連絡先フォルダーから連絡先を取得します。");
        command.Options.Add(name);
        command.Options.Add(company);
        command.SetAction(result => ExecuteList(result.GetValue(name), result.GetValue(company)));
        return command;
    }

    private Command CreateShow()
    {
        var entryId = new Option<string>("--entry-id")
        {
            Description = "連絡先の Entry ID です。",
            Required = true,
        };
        var command = new Command("show", "指定した連絡先の詳細を取得します。");
        command.Options.Add(entryId);
        command.SetAction(result => ExecuteShow(result.GetValue(entryId)));
        return command;
    }

    private Command CreateAdd()
    {
        var name = new Option<string>("--name")
        {
            Description = "連絡先名です。",
            Required = true,
        };
        var email = new Option<string>("--email")
        {
            Description = "メールアドレスです。",
            Required = true,
        };
        var company = new Option<string>("--company") { Description = "会社名です。" };
        var command = new Command("add", "既定の連絡先フォルダーに連絡先を追加します。");
        command.Options.Add(name);
        command.Options.Add(email);
        command.Options.Add(company);
        command.SetAction(result => ExecuteAdd(
            result.GetRequiredValue(name),
            result.GetRequiredValue(email),
            result.GetValue(company)));
        return command;
    }

    private Command CreateLookup()
    {
        var name = new Option<string>("--name")
        {
            Description = "Exchange 社内アドレス帳で解決する名前です。",
            Required = true,
        };
        var command = new Command("lookup", "Exchange 社内アドレス帳から名前を解決します。");
        command.Options.Add(name);
        command.SetAction(result => ExecuteLookup(result.GetValue(name)));
        return command;
    }

    private int ExecuteList(string? name, string? company)
    {
        return CommandExecution.RunList(error, "連絡先一覧の取得", () => service.List(name, company),
            WriteSummary);
    }

    private int ExecuteShow(string? entryId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
        {
            error.WriteLine("--entry-id に連絡先の Entry ID を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "連絡先の取得", () =>
        {
            var contact = service.Get(entryId!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                contact.EntryId,
                contact.Name,
                contact.Email,
                contact.Email2,
                contact.Email3,
                contact.Company,
                contact.BusinessPhone,
            });
        });
    }

    private int ExecuteAdd(string name, string email, string? company)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            error.WriteLine("--name に連絡先名を指定してください。");
            return 1;
        }

        if (string.IsNullOrWhiteSpace(email))
        {
            error.WriteLine("--email にメールアドレスを指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "連絡先の追加", () =>
            WriteSummary(service.Add(name, email, company)));
    }

    private int ExecuteLookup(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            error.WriteLine("--name に名前を指定してください。");
            return 1;
        }

        return CommandExecution.Run(error, "Exchange 社内アドレス帳の検索", () =>
        {
            var contact = service.Lookup(name!);
            json.WriteLine(new
            {
                schemaVersion = 1,
                contact.Name,
                contact.Email,
                contact.Company,
            });
        });
    }

    private void WriteSummary(Contact contact)
    {
        json.WriteLine(new
        {
            schemaVersion = 1,
            contact.EntryId,
            contact.Name,
            contact.Email,
            contact.Company,
        });
    }
}
