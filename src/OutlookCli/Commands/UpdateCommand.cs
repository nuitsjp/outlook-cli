using System.CommandLine;
using System.ComponentModel;
using System.Net.Http;
using Microsoft.Win32;
using OutlookCli.Updates;

namespace OutlookCli.Commands;

internal sealed class UpdateCommand(JsonOutput json, TextWriter error)
{
    public Command Create()
    {
        var command = new Command("update", "GitHub Releases からインストール済みの CLI を更新します。");
        command.SetAction((_, cancellationToken) => ExecuteAsync(cancellationToken));
        return command;
    }

    private async Task<int> ExecuteAsync(CancellationToken cancellationToken)
    {
        try
        {
            var assembly = typeof(UpdateCommand).Assembly;
            using var registry = Registry.CurrentUser.OpenSubKey(@"Software\OutlookCli\Shim");
            if (registry?.GetValue("Target") is not string target
                || !string.Equals(Path.GetFullPath(target), assembly.Location, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("インストーラーでインストールした outlook-cli から update を実行してください。");
            }

            using var http = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false, UseCookies = false });
            var updater = new ReleaseUpdater(http, new UpdateInstaller().Start);
            var result = await updater.UpdateAsync(
                assembly.GetName().Version!, Path.GetDirectoryName(assembly.Location)!, cancellationToken);
            json.WriteLine(new
            {
                schemaVersion = 1,
                result.Status,
                result.CurrentVersion,
                result.LatestVersion,
                result.LogPath,
            });
            return 0;
        }
        catch (HttpRequestException exception)
        {
            error.WriteLine($"GitHub Releases へのアクセスに失敗しました。{exception.Message}");
            return 1;
        }
        catch (OperationCanceledException)
        {
            error.WriteLine("更新を中断しました。");
            return 1;
        }
        catch (Exception exception) when (exception is InvalidOperationException or IOException
            or UnauthorizedAccessException or Win32Exception or FormatException)
        {
            error.WriteLine($"更新を準備できませんでした。{exception.Message}");
            return 1;
        }
    }
}
