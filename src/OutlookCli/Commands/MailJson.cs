using System.Text.Json;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal static class MailJson
{
    public static JsonSerializerOptions Options { get; } = new(JsonSerializerDefaults.Web);

    public static void WriteSummary(TextWriter output, MailSummary mail)
    {
        output.WriteLine(JsonSerializer.Serialize(new
        {
            schemaVersion = 1,
            mail.EntryId,
            mail.ReceivedAt,
            mail.SenderName,
            mail.Subject,
        }, Options));
    }
}
