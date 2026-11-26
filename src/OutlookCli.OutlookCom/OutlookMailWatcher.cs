using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using OutlookCli.Core.Mail;

namespace OutlookCli.OutlookCom;

public sealed class OutlookMailWatcher : IMailWatcher
{
    private static readonly Guid ApplicationEvents = new("0006302C-0000-0000-C000-000000000046");
    private const int NewMailEx = 0xFAB5; // ApplicationEvents_11.NewMailEx

    public void Watch(Action<MailSummary> onMail, Action onReady, CancellationToken cancellationToken)
    {
        using var inbox = OutlookInbox.Open();
        using var stopped = new ManualResetEvent(false);
        using var cancellation = cancellationToken.Register(() => stopped.Set());
        ExceptionDispatchInfo? failure = null;

        void OnReceived(string entryId)
        {
            if (cancellationToken.IsCancellationRequested || failure is not null)
            {
                return;
            }

            try
            {
                var item = inbox.GetReceivedItem(entryId);
                try
                {
                    if (item is not null && (int)((dynamic)item).Class == OutlookInbox.MailItemClass)
                    {
                        onMail(OutlookMailQuery.ReadSummary(item));
                    }
                }
                finally
                {
                    OutlookInbox.Release(item);
                }
            }
            catch (Exception exception)
            {
                failure = ExceptionDispatchInfo.Capture(exception);
                stopped.Set();
            }
        }

        Action<string> handler = OnReceived;
        ComEventsHelper.Combine(inbox.Application, ApplicationEvents, NewMailEx, handler);
        try
        {
            onReady();
            // Wait indefinitely on the STA while COM dispatches incoming event calls.
            var handle = stopped.SafeWaitHandle.DangerousGetHandle();
            Marshal.ThrowExceptionForHR(CoWaitForMultipleHandles(
                0, uint.MaxValue, 1, in handle, out _));
            failure?.Throw();
        }
        finally
        {
            ComEventsHelper.Remove(inbox.Application, ApplicationEvents, NewMailEx, handler);
        }
    }

    [DllImport("ole32.dll")]
    private static extern int CoWaitForMultipleHandles(
        uint flags, uint timeout, uint count, in nint handles, out uint index);
}
