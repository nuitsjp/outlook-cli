using OutlookCli.Core.Contacts;

namespace OutlookCli.OutlookCom;

public sealed class OutlookContactService : IContactService
{
    private const int ContactsFolder = 10; // OlDefaultFolders.olFolderContacts
    private const int ContactItemClass = 40; // OlObjectClass.olContact

    public IEnumerable<Contact> List(string? name, string? company)
    {
        using var contacts = OutlookFolderConnection.Open(ContactsFolder);
        object? items = null;
        try
        {
            items = ((dynamic)contacts.Folder).Items;
            var results = (dynamic)items!;
            results.Sort("[FullName]", false);

            for (object? item = results.GetFirst(); item is not null; item = results.GetNext())
            {
                Contact contact;
                try
                {
                    if ((int)((dynamic)item).Class != ContactItemClass)
                    {
                        continue;
                    }

                    contact = ReadContact(item);
                }
                finally
                {
                    OutlookFolderConnection.Release(item);
                }

                if (!Matches(contact.Name, name) || !Matches(contact.Company, company))
                {
                    continue;
                }

                yield return contact;
            }
        }
        finally
        {
            OutlookFolderConnection.Release(items);
        }
    }

    public Contact Get(string entryId)
    {
        using var contacts = OutlookFolderConnection.Open(ContactsFolder);
        object? item = null;
        try
        {
            item = contacts.GetItem(entryId);
            if (item is null || (int)((dynamic)item).Class != ContactItemClass)
            {
                throw new InvalidOperationException("指定された ID に対応する連絡先が見つかりません。");
            }

            EnsureInContactsFolder(item, contacts.Folder);
            return ReadContact(item);
        }
        finally
        {
            OutlookFolderConnection.Release(item);
        }
    }

    public Contact Add(string name, string email, string? company)
    {
        using var contacts = OutlookFolderConnection.Open(ContactsFolder);
        object? item = null;
        object? items = null;
        try
        {
            items = ((dynamic)contacts.Folder).Items;
            item = ((dynamic)items).Add("IPM.Contact");

            var contact = (dynamic)item!;
            contact.FullName = name;
            contact.Email1AddressType = "SMTP";
            contact.Email1Address = email;
            if (company is not null)
            {
                contact.CompanyName = company;
            }

            contact.Save();
            return ReadContact(item);
        }
        finally
        {
            OutlookFolderConnection.Release(items);
            OutlookFolderConnection.Release(item);
        }
    }

    public Contact Lookup(string name)
    {
        // ResolveRecipient uses NameSpace.CreateRecipient followed by Recipient.Resolve,
        // which is Outlook's documented name-resolution path for the address book.
        using var contacts = OutlookFolderConnection.Open();
        object? recipient = null;
        object? addressEntry = null;
        object? exchangeObject = null;
        try
        {
            // OutlookFolderConnection owns the MAPI session and performs the
            // documented CreateRecipient/Recipient.Resolve sequence.
            recipient = contacts.ResolveRecipient(name);
            if (recipient is null)
            {
                throw new InvalidOperationException("指定された名前を Exchange 社内アドレス帳で解決できませんでした。");
            }

            addressEntry = ((dynamic)recipient).AddressEntry;
            if (addressEntry is null)
            {
                throw new InvalidOperationException("指定された名前を Exchange 社内アドレス帳で解決できませんでした。");
            }

            var addressEntryType = (int)((dynamic)addressEntry).AddressEntryUserType;
            var entryId = (string?)((dynamic)addressEntry).ID
                ?? (string?)((dynamic)recipient).EntryID
                ?? string.Empty;

            switch (addressEntryType)
            {
                case 0: // olExchangeUserAddressEntry
                case 5: // olExchangeRemoteUserAddressEntry
                    exchangeObject = ((dynamic)addressEntry).GetExchangeUser();
                    if (exchangeObject is null)
                    {
                        throw new InvalidOperationException("指定された名前を Exchange 社内アドレス帳で解決できませんでした。");
                    }

                    return new Contact(
                        entryId,
                        (string?)((dynamic)exchangeObject).Name
                            ?? (string?)((dynamic)addressEntry).Name
                            ?? string.Empty,
                        (string?)((dynamic)exchangeObject).PrimarySmtpAddress
                            ?? string.Empty,
                        (string?)((dynamic)exchangeObject).CompanyName
                            ?? string.Empty,
                        BusinessPhone: (string?)((dynamic)exchangeObject).BusinessTelephoneNumber
                            ?? string.Empty);

                case 1: // olExchangeDistributionListAddressEntry
                    exchangeObject = ((dynamic)addressEntry).GetExchangeDistributionList();
                    if (exchangeObject is null)
                    {
                        throw new InvalidOperationException("指定された名前を Exchange 社内アドレス帳で解決できませんでした。");
                    }

                    return new Contact(
                        entryId,
                        (string?)((dynamic)exchangeObject).Name
                            ?? (string?)((dynamic)addressEntry).Name
                            ?? string.Empty,
                        (string?)((dynamic)exchangeObject).PrimarySmtpAddress
                            ?? string.Empty,
                        string.Empty);

                default:
                    throw new InvalidOperationException("指定された名前は Exchange 社内アドレス帳の連絡先ではありません。");
            }
        }
        finally
        {
            OutlookFolderConnection.Release(exchangeObject);
            OutlookFolderConnection.Release(addressEntry);
            OutlookFolderConnection.Release(recipient);
        }
    }

    private static Contact ReadContact(object item)
    {
        var contact = (dynamic)item;
        return new Contact(
            (string?)contact.EntryID ?? string.Empty,
            (string?)contact.FullName ?? string.Empty,
            (string?)contact.Email1Address ?? string.Empty,
            (string?)contact.CompanyName ?? string.Empty,
            (string?)contact.Email2Address ?? string.Empty,
            (string?)contact.Email3Address ?? string.Empty,
            (string?)contact.BusinessTelephoneNumber ?? string.Empty);
    }

    private static bool Matches(string value, string? filter) =>
        string.IsNullOrWhiteSpace(filter)
            || value.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0;

    private static void EnsureInContactsFolder(object item, object folder)
    {
        object? parent = null;
        try
        {
            parent = ((dynamic)item).Parent;
            if (!string.Equals(
                (string)((dynamic)parent!).EntryID,
                (string)((dynamic)folder).EntryID,
                StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("指定された連絡先は既定の連絡先フォルダーにありません。");
            }
        }
        finally
        {
            OutlookFolderConnection.Release(parent);
        }
    }
}
