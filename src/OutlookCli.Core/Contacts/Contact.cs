namespace OutlookCli.Core.Contacts;

public sealed record Contact(
    string EntryId,
    string Name,
    string Email,
    string Company,
    string Email2 = "",
    string Email3 = "",
    string BusinessPhone = "");
