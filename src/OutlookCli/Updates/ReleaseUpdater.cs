using System.Globalization;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OutlookCli.Updates;

internal sealed class ReleaseUpdater(HttpClient httpClient, Func<string, string, string> startInstaller)
{
    private const string ReleasesUrl = "https://api.github.com/repos/nuitsjp/outlook-cli/releases";

    public async Task<UpdateResult> UpdateAsync(
        Version installedVersion, string installDirectory, CancellationToken cancellationToken)
    {
        var current = new Version(installedVersion.Major, installedVersion.Minor, installedVersion.Build);
        using var releaseRequest = CreateApiRequest(ReleasesUrl + "/latest", "application/vnd.github+json");
        using var releaseResponse = await httpClient.SendAsync(releaseRequest, cancellationToken);
        releaseResponse.EnsureSuccessStatusCode();

        Release? release;
        try
        {
            release = JsonSerializer.Deserialize<Release>(await releaseResponse.Content.ReadAsStringAsync());
        }
        catch (JsonException exception)
        {
            throw new InvalidOperationException("GitHub Releases の応答を読み取れませんでした。", exception);
        }
        if (release?.Tag is not string tag
            || release.Draft
            || release.Prerelease
            || !tag.StartsWith("v", StringComparison.Ordinal)
            || !Version.TryParse(tag.Substring(1), out var latest)
            || latest.Build < 0
            || latest.Revision >= 0
            || tag != "v" + latest.ToString(3))
        {
            throw new InvalidOperationException("GitHub Releases の正式リリースを確認できませんでした。");
        }

        var currentText = current.ToString(3);
        var latestText = latest.ToString(3);
        if (latest <= current)
        {
            return new UpdateResult(currentText, latestText, "up-to-date", null);
        }

        var installerName = $"outlook-cli-{latestText}-setup.exe";
        var asset = release.Assets?.SingleOrDefault(item => item.Name == installerName);
        if (asset is null || asset.Id <= 0)
        {
            throw new InvalidOperationException($"GitHub Releases に {installerName} が見つかりません。");
        }

        const string digestPrefix = "sha256:";
        var digest = asset.Digest;
        if (digest is null
            || !digest.StartsWith(digestPrefix, StringComparison.Ordinal)
            || digest.Length != digestPrefix.Length + 64
            || !digest.Skip(digestPrefix.Length).All(Uri.IsHexDigit))
        {
            throw new InvalidOperationException("インストーラーの SHA-256 を GitHub Releases から取得できませんでした。");
        }

        var tempDirectory = Path.Combine(Path.GetTempPath(), "outlook-cli-update-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDirectory);
        var installerPath = Path.Combine(tempDirectory, installerName);
        var handedOff = false;
        try
        {
            await DownloadAsync(asset.Id, installerPath, cancellationToken);
            using (var installer = File.OpenRead(installerPath))
            using (var sha256 = SHA256.Create())
            {
                var actual = BitConverter.ToString(sha256.ComputeHash(installer)).Replace("-", string.Empty);
                if (!string.Equals(actual, digest.Substring(digestPrefix.Length), StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException("インストーラーの SHA-256 が GitHub Releases の値と一致しません。");
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            var logPath = startInstaller(installerPath, installDirectory);
            handedOff = true;
            return new UpdateResult(currentText, latestText, "scheduled", logPath);
        }
        finally
        {
            if (!handedOff)
            {
                File.Delete(installerPath);
                Directory.Delete(tempDirectory);
            }
        }
    }

    private async Task DownloadAsync(long assetId, string path, CancellationToken cancellationToken)
    {
        var address = ReleasesUrl + "/assets/" + assetId.ToString(CultureInfo.InvariantCulture);
        using var apiRequest = CreateApiRequest(address, "application/octet-stream");
        using var apiResponse = await httpClient.SendAsync(apiRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
        if (apiResponse.StatusCode == HttpStatusCode.Redirect)
        {
            var downloadUrl = apiResponse.Headers.Location;
            if (downloadUrl is not { IsAbsoluteUri: true, Scheme: "https" })
            {
                throw new InvalidOperationException("GitHub Releases から HTTPS のダウンロード先を取得できませんでした。");
            }

            using var downloadRequest = new HttpRequestMessage(HttpMethod.Get, downloadUrl);
            using var download = await httpClient.SendAsync(downloadRequest, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            await SaveAsync(download, path, cancellationToken);
            return;
        }

        await SaveAsync(apiResponse, path, cancellationToken);
    }

    private static async Task SaveAsync(HttpResponseMessage response, string path, CancellationToken cancellationToken)
    {
        response.EnsureSuccessStatusCode();
        using var source = await response.Content.ReadAsStreamAsync();
        using var destination = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        await source.CopyToAsync(destination, 81920, cancellationToken);
    }

    private static HttpRequestMessage CreateApiRequest(string address, string accept)
    {
        var request = new HttpRequestMessage(HttpMethod.Get, address);
        request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue(accept));
        request.Headers.UserAgent.ParseAdd("outlook-cli");
        request.Headers.Add("X-GitHub-Api-Version", "2022-11-28");
        return request;
    }

    private sealed class Release
    {
        [JsonPropertyName("tag_name")]
        public string? Tag { get; set; }

        [JsonPropertyName("draft")]
        public bool Draft { get; set; }

        [JsonPropertyName("prerelease")]
        public bool Prerelease { get; set; }

        [JsonPropertyName("assets")]
        public Asset[]? Assets { get; set; }
    }

    private sealed class Asset
    {
        [JsonPropertyName("id")]
        public long Id { get; set; }

        [JsonPropertyName("name")]
        public string? Name { get; set; }

        [JsonPropertyName("digest")]
        public string? Digest { get; set; }
    }
}

internal sealed class UpdateResult(string currentVersion, string latestVersion, string status, string? logPath)
{
    public string CurrentVersion { get; } = currentVersion;
    public string LatestVersion { get; } = latestVersion;
    public string Status { get; } = status;
    public string? LogPath { get; } = logPath;
}
