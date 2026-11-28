using System.Runtime.InteropServices;
using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class OutlookMailSenderTests
{
    [Fact]
    public void Send_marks_submitted_after_send_and_attempts_all_releases_in_order()
    {
        var application = new MailApplicationDouble();
        var releaseOrder = new List<object>();
        var cleanup = CreateCleanup(
            releaseOrder,
            application.Mail.Recipients.Recipient,
            application.Mail);

        OutlookMailSender.Send(cleanup, application, "to@example.com", "subject", "body");

        Assert.Equal(0, application.CreatedItemType);
        Assert.True(application.Mail.SendCalled);
        Assert.True(cleanup.MailSubmitted);
        Assert.Equal(
            [application.Mail.Recipients.Recipient, application.Mail.Recipients, application.Mail],
            releaseOrder);
        Assert.Equal(2, cleanup.Failures.Count);
        Assert.Equal("to@example.com", application.Mail.Recipients.Address);
        Assert.Equal("subject", application.Mail.Subject);
        Assert.Equal(1, application.Mail.BodyFormat);
        Assert.Equal("body", application.Mail.Body);
    }

    [Fact]
    public void Send_preserves_send_com_exception_and_releases_remaining_references()
    {
        var operation = new COMException("Send failed", unchecked((int)0x80004005));
        var application = new MailApplicationDouble();
        application.Mail.SendException = operation;
        var releaseOrder = new List<object>();
        var cleanup = CreateCleanup(
            releaseOrder,
            application.Mail.Recipients.Recipient,
            application.Mail.Recipients);

        var exception = Assert.Throws<COMException>(() =>
            OutlookMailSender.Send(cleanup, application, "to@example.com", "subject", "body"));

        Assert.Same(operation, exception);
        Assert.False(cleanup.MailSubmitted);
        Assert.Equal(
            [application.Mail.Recipients.Recipient, application.Mail.Recipients, application.Mail],
            releaseOrder);
        Assert.Equal(2, cleanup.Failures.Count);
    }

    private static ComCleanup CreateCleanup(List<object> releaseOrder, params object[] failures)
    {
        return new ComCleanup(value =>
        {
            releaseOrder.Add(value);
            if (failures.Contains(value))
            {
                throw new InvalidOperationException($"release failed: {value.GetType().Name}");
            }
        });
    }
}

public sealed class MailApplicationDouble
{
    public MailItemDouble Mail { get; } = new();

    public int CreatedItemType { get; private set; }

    public MailItemDouble CreateItem(int itemType)
    {
        CreatedItemType = itemType;
        return Mail;
    }
}

public sealed class MailItemDouble
{
    public MailRecipientsDouble Recipients { get; } = new();

    public string? Subject { get; set; }

    public int BodyFormat { get; set; }

    public string? Body { get; set; }

    public Exception? SendException { get; set; }

    public bool SendCalled { get; private set; }

    public void Send()
    {
        SendCalled = true;
        if (SendException is not null)
        {
            throw SendException;
        }
    }
}

public sealed class MailRecipientsDouble
{
    public MailRecipientDouble Recipient { get; } = new();

    public string? Address { get; private set; }

    public MailRecipientDouble Add(string address)
    {
        Address = address;
        return Recipient;
    }
}

public sealed class MailRecipientDouble
{
    public bool Resolve() => true;
}
