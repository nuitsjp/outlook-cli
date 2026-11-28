using System.Runtime.InteropServices;
using OutlookCli.OutlookCom;
using Xunit;

namespace OutlookCli.Tests;

public sealed class AttachmentFileTests
{
    [Fact]
    public void Save_moves_the_downloaded_file_without_leaving_a_temporary_file()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");

        AttachmentFileSaver.Save(destination, temporaryPath =>
            File.WriteAllText(temporaryPath, "downloaded"));

        Assert.Equal("downloaded", File.ReadAllText(destination));
        Assert.Equal([destination], Directory.GetFiles(directory.Path));
    }

    [Fact]
    public void Save_rejects_an_existing_destination_without_calling_the_writer()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");
        File.WriteAllText(destination, "other process");
        var called = false;

        Assert.Throws<InvalidOperationException>(() => AttachmentFileSaver.Save(
            destination,
            _ => called = true));

        Assert.False(called);
        Assert.Equal("other process", File.ReadAllText(destination));
    }

    [Fact]
    public void Save_removes_the_temporary_file_when_the_writer_fails()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");
        var temporaryPath = Path.Combine(directory.Path, "download.tmp");

        Assert.Throws<InvalidOperationException>(() => AttachmentFileSaver.Save(
            destination,
            path =>
            {
                Assert.Equal(temporaryPath, path);
                File.WriteAllText(path, "partial");
                throw new InvalidOperationException("download failed");
            },
            _ => temporaryPath));

        Assert.False(File.Exists(temporaryPath));
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public void Save_does_not_truncate_a_colliding_temporary_candidate()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");
        var occupied = Path.Combine(directory.Path, "occupied.tmp");
        var available = Path.Combine(directory.Path, "available.tmp");
        File.WriteAllText(occupied, "keep this file");
        var candidates = new Queue<string>([occupied, available]);

        AttachmentFileSaver.Save(
            destination,
            path => File.WriteAllText(path, "downloaded"),
            _ => candidates.Dequeue());

        Assert.Equal("keep this file", File.ReadAllText(occupied));
        Assert.Equal("downloaded", File.ReadAllText(destination));
        Assert.False(File.Exists(available));
    }

    [Fact]
    public void Temporary_file_cannot_be_the_destination_itself()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "report.tmp");
        var separate = Path.Combine(directory.Path, "partial.tmp");
        var candidates = new Queue<string>([destination, separate]);

        AttachmentFileSaver.Save(destination, path =>
        {
            Assert.False(File.Exists(destination));
            File.WriteAllText(path, "complete");
        }, _ => candidates.Dequeue());

        Assert.Equal("complete", File.ReadAllText(destination));
        Assert.False(File.Exists(separate));
    }

    [Fact]
    public void Save_preserves_a_destination_created_after_the_initial_check()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");
        var temporaryPath = Path.Combine(directory.Path, "download.tmp");

        Assert.Throws<InvalidOperationException>(() => AttachmentFileSaver.Save(
            destination,
            path =>
            {
                File.WriteAllText(path, "downloaded");
                File.WriteAllText(destination, "created by another process");
            },
            _ => temporaryPath));

        Assert.Equal("created by another process", File.ReadAllText(destination));
        Assert.False(File.Exists(temporaryPath));
    }

    [Fact]
    public void Save_converts_a_temporary_directory_failure_to_an_operation_error()
    {
        using var directory = TemporaryDirectory.Create();
        var missingDirectory = Path.Combine(directory.Path, "missing");
        var destination = Path.Combine(missingDirectory, "attachment.bin");

        var exception = Assert.Throws<InvalidOperationException>(() => AttachmentFileSaver.Save(
            destination,
            _ => throw new InvalidOperationException("writer must not run")));

        Assert.IsType<DirectoryNotFoundException>(exception.InnerException);
    }

    [Fact]
    public void Save_preserves_the_outlook_failure_when_the_temporary_file_is_locked()
    {
        using var directory = TemporaryDirectory.Create();
        var destination = Path.Combine(directory.Path, "attachment.bin");
        var primary = new COMException("SaveAsFile failed", unchecked((int)0x80004005));
        FileStream? lockFile = null;
        try
        {
            var error = Assert.Throws<OutlookComCleanupException>(() =>
                AttachmentFileSaver.Save(destination, path =>
                {
                    lockFile = File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    throw primary;
                }));

            Assert.Same(primary, error.PrimaryException);
            Assert.IsType<IOException>(Assert.Single(error.CleanupExceptions).InnerException);
            Assert.False(File.Exists(destination));
        }
        finally
        {
            lockFile?.Dispose();
        }
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        private TemporaryDirectory(string path) => Path = path;

        public string Path { get; }

        public static TemporaryDirectory Create()
        {
            var path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "outlook-cli-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return new TemporaryDirectory(path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
