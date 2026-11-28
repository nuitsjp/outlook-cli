using System.Text.Json.Serialization;
using OutlookCli.Core.Mail;

namespace OutlookCli.Commands;

internal class MailSummaryJson(MailSummary summary)
{
    public int SchemaVersion => 1;
    public string EntryId => summary.EntryId;
    public DateTimeOffset ReceivedAt => summary.ReceivedAt;
    public string SenderName => summary.SenderName;
    public string Subject => summary.Subject;
}

internal sealed class MailDetailJson(MailDetail detail) : MailSummaryJson(detail.Summary)
{
    [JsonPropertyOrder(1)] public string To => detail.To;
    [JsonPropertyOrder(1)] public string Cc => detail.Cc;
    [JsonPropertyOrder(1)] public string Bcc => detail.Bcc;
    [JsonPropertyOrder(1)] public string Body => detail.Body;
}
