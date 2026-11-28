using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal static class MailJson
{
    public static void WriteSummary(JsonOutput json, MailSummary mail)
    {
        json.WriteLine(new
        {
            schemaVersion = 1,
            mail.EntryId,
            mail.ReceivedAt,
            mail.SenderName,
            mail.Subject,
        });
    }
}
