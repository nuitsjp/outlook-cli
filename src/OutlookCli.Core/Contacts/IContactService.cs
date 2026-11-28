namespace OutlookCli.Core.Contacts;

public interface IContactService
{
    IEnumerable<Contact> List(string? name, string? company);

    Contact Get(string entryId);

    Contact Add(string name, string email, string? company);

    Contact Lookup(string name);
}
