using System.Text.Json.Serialization;
using OutlookCli.Core.Calendar;

namespace OutlookCli.Commands;

internal class CalendarSummaryJson(CalendarSummary summary)
{
    public int SchemaVersion => 1;
    public string Id => summary.Id;
    public string Subject => summary.Subject;
    public DateTimeOffset StartAt => summary.StartAt;
    public DateTimeOffset EndAt => summary.EndAt;
    public bool IsAllDay => summary.IsAllDay;
    public string Location => summary.Location;
}

internal sealed class CalendarDetailJson(CalendarDetail detail) : CalendarSummaryJson(detail.Summary)
{
    [JsonPropertyOrder(1), Newtonsoft.Json.JsonProperty(Order = 1)]
    public string Organizer => detail.Organizer;
    [JsonPropertyOrder(1), Newtonsoft.Json.JsonProperty(Order = 1)]
    public IReadOnlyList<CalendarAttendee> Attendees => detail.Attendees;
    [JsonPropertyOrder(1), Newtonsoft.Json.JsonProperty(Order = 1)]
    public string Body => detail.Body;
}
