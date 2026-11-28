namespace OutlookCli.OutlookCom;

internal static class AttachmentFileSaver
{
    internal static void Save(
        string destinationPath,
        Action<string> saveToPath,
        Func<string, string>? temporaryPathFactory = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        ArgumentNullException.ThrowIfNull(saveToPath);

        var fullPath = Path.GetFullPath(destinationPath);
        if (File.Exists(fullPath))
        {
            throw new InvalidOperationException($"保存先ファイルが既に存在します。上書きしません: {fullPath}");
        }

        var directory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        var factory = temporaryPathFactory ?? (folder => Path.Combine(folder, Path.GetRandomFileName()));
        var temporaryPath = ReserveTemporaryPath(directory, fullPath, factory);
        try
        {
            saveToPath(temporaryPath);
            try
            {
                // File.Move without overwrite is the final race-safe destination check.
                File.Move(temporaryPath, fullPath, overwrite: false);
            }
            catch (IOException exception)
            {
                throw new InvalidOperationException(
                    $"保存先ファイルが既に存在するか、保存できません。上書きしません: {fullPath}",
                    exception);
            }
            catch (UnauthorizedAccessException exception)
            {
                throw new InvalidOperationException(
                    $"保存先ファイルへ書き込めません。上書きしません: {fullPath}",
                    exception);
            }
        }
        catch (Exception primary)
        {
            // After a successful move the temporary name can belong to another writer.
            try
            {
                File.Delete(temporaryPath);
            }
            catch (Exception failure) when (failure is IOException or UnauthorizedAccessException)
            {
                var cleanup = new InvalidOperationException($"一時ファイルを削除できません: {temporaryPath}", failure);
                throw OutlookComCleanupException.Combine(primary, [cleanup], mailSubmitted: false);
            }

            throw;
        }
    }

    private static string ReserveTemporaryPath(string directory, string destination, Func<string, string> factory)
    {
        while (true)
        {
            var path = factory(directory);
            if (string.Equals(path, destination, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                // CreateNew prevents a stale/colliding candidate from being truncated.
                using (new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                }

                return path;
            }
            catch (IOException) when (File.Exists(path))
            {
                // A random candidate collided; choose another one without touching it.
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                throw new InvalidOperationException(
                    $"一時ファイルを作成できません。保存先を確認してください: {directory}",
                    exception);
            }
        }
    }
}
