using System.Diagnostics;
using System.Runtime.InteropServices;
using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class OutlookFolderConnectionTests
{
    [Fact]
    public void Failed_shared_folder_lookup_releases_every_acquired_reference()
    {
        var lookupFailure = new COMException("GetSharedDefaultFolder failed", unchecked((int)0x80070005));
        var releaseFailure = new InvalidOperationException("recipient release failed");
        var session = new FakeSession(lookupFailure);
        var instance = new FakeApplication(session);
        var released = new List<object>();
        var cleanup = new ComCleanup(value =>
        {
            released.Add(value);
            if (ReferenceEquals(value, session.Recipient))
            {
                throw releaseFailure;
            }
        });
        var application = new OutlookApplication(cleanup, instance, null);

        var error = Assert.Throws<COMException>(() =>
            OutlookFolderConnection.Open(cleanup, application, 9, "calendar owner"));

        Assert.Same(lookupFailure, error);
        Assert.Same(releaseFailure, Assert.Single(cleanup.Failures));
        Assert.Collection(released,
            value => Assert.Same(session.Recipient, value),
            value => Assert.Same(session, value),
            value => Assert.Same(instance, value));
    }

    [Fact]
    public void Failed_folder_release_still_attempts_session_application_and_shutdown()
    {
        var operation = new COMException("GetDefaultFolder failed", unchecked((int)0x80004005));
        var failure = Assert.Throws<OutlookComCleanupException>(() => ComCleanup.Run(cleanup =>
        {
            // Each non-COM reference fails to release. A process without a PID fails to wait.
            var application = new OutlookApplication(cleanup, new object(), new Process());
            using var connection = new OutlookFolderConnection(cleanup, application, new object(), new object());
            throw operation;
        }));

        Assert.Same(operation, failure.PrimaryException);
        Assert.Equal(operation.HResult, failure.HResult);
        Assert.Collection(failure.CleanupExceptions,
            error => Assert.IsType<ArgumentException>(error),
            error => Assert.IsType<ArgumentException>(error),
            error => Assert.IsType<ArgumentException>(error),
            error => Assert.IsType<InvalidOperationException>(error));
    }

    public sealed class FakeApplication(FakeSession session)
    {
        public FakeSession GetNamespace(string name) => session;
    }

    public sealed class FakeSession(COMException failure)
    {
        public FakeRecipient Recipient { get; } = new();
        public FakeRecipient CreateRecipient(string name) => Recipient;
        public object GetSharedDefaultFolder(FakeRecipient recipient, int folder) => throw failure;
    }

    public sealed class FakeRecipient
    {
        public bool Resolve() => true;
    }
}
