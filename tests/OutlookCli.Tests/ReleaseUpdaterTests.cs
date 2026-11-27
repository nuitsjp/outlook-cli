using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OutlookCli.Updates;
using Xunit;

namespace OutlookCli.Tests;

public sealed class ReleaseUpdaterTests
{
    private const string LatestUrl =
        "https://api.github.com/repos/nuitsjp/outlook-cli/releases/latest";
    private const string AssetUrl =
        "https://api.github.com/repos/nuitsjp/outlook-cli/releases/assets/123";

    [Fact]
    public async Task Newer_release_downloads_verified_installer_and_hands_it_to_worker()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var digest = Sha256Digest(payload);
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var downloadUrl = "https://downloads.example.test/" + installerName;
        var events = new List<string>();
        var handler = new RecordingHandler(request =>
        {
            events.Add(request.RequestUri!.AbsoluteUri);
            if (request.RequestUri!.AbsoluteUri == LatestUrl)
            {
                AssertApiRequest(request, "application/vnd.github+json");
                return Reply(JsonResponse(ReleaseJson(
                    "1.2.4",
                    new object[]
                    {
                        Asset(123, installerName, digest),
                        Asset(124, "outlook-cli-1.2.4-setup.exe.asc", digest),
                    })));
            }

            if (request.RequestUri.AbsoluteUri == AssetUrl)
            {
                AssertApiRequest(request, "application/octet-stream");
                return Reply(RedirectResponse(downloadUrl));
            }

            if (request.RequestUri.AbsoluteUri == downloadUrl)
            {
                Assert.Null(request.Headers.Authorization);
                return Reply(BinaryResponse(payload));
            }

            return Failure<HttpResponseMessage>("Unexpected request: " + request.RequestUri);
        });

        using var client = new HttpClient(handler);
        string? handedOffPath = null;
        string? handedOffInstallDirectory = null;
        var confirmations = new List<(string Current, string Latest)>();
        var logPath = Path.Combine(fixture.Path, "update.log");
        var updater = new ReleaseUpdater(client, (path, installDirectory) =>
        {
            handedOffPath = path;
            handedOffInstallDirectory = installDirectory;
            Assert.True(File.Exists(path));
            Assert.Equal(payload, File.ReadAllBytes(path));

            // The staged path is intentionally left for the handed-off
            // installer process; this test only cleans up its own fixture.
            return logPath;
        }, (current, latest) =>
        {
            confirmations.Add((current, latest));
            events.Add("confirm");
            return true;
        });

        var result = await updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None);

        Assert.Equal("1.2.3", result.CurrentVersion);
        Assert.Equal("1.2.4", result.LatestVersion);
        Assert.Equal("scheduled", result.Status);
        Assert.Equal(logPath, result.LogPath);
        Assert.NotNull(handedOffPath);
        Assert.Equal(fixture.Path, handedOffInstallDirectory);
        Assert.Equal([("1.2.3", "1.2.4")], confirmations);
        Assert.Equal([LatestUrl, "confirm", AssetUrl, downloadUrl], events);

        Assert.Equal(
            [LatestUrl, AssetUrl, downloadUrl],
            handler.Requests.Select(request => request.Uri.AbsoluteUri).ToArray());
        Assert.All(handler.Requests, request =>
        {
            Assert.Null(request.Authorization);
        });
    }

    [Fact]
    public async Task Declined_update_keeps_versions_without_downloading_or_handoff()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            if (request.RequestUri.AbsoluteUri == LatestUrl)
            {
                return Reply(JsonResponse(ReleaseJson(
                    "1.2.4", new object[] { Asset(123, installerName, Sha256Digest(payload)) })));
            }

            return Failure<HttpResponseMessage>("A declined update must not download the asset");
        });
        using var client = new HttpClient(handler);
        var handoffCount = 0;
        var updater = new ReleaseUpdater(client, (_, _) =>
        {
            handoffCount++;
            return "must-not-start";
        }, (_, _) => false);

        var result = await updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None);

        Assert.Equal("1.2.3", result.CurrentVersion);
        Assert.Equal("1.2.4", result.LatestVersion);
        Assert.Equal("declined", result.Status);
        Assert.Null(result.LogPath);
        Assert.Equal(0, handoffCount);
        Assert.Single(handler.Requests);
    }

    [Theory]
    [InlineData("1.2.4")]
    [InlineData("1.2.5")]
    public async Task Equal_or_newer_release_is_up_to_date_without_assets_or_handoff(string installedVersion)
    {
        using var fixture = new OwnedDirectory();
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(ReleaseJson("1.2.4", assets: null)));
        });
        using var client = new HttpClient(handler);
        var handoffCount = 0;
        var confirmationCount = 0;
        var updater = new ReleaseUpdater(client, (_, _) =>
        {
            handoffCount++;
            return "must-not-start";
        }, (_, _) =>
        {
            confirmationCount++;
            return true;
        });

        var result = await updater.UpdateAsync(
            Version.Parse(installedVersion), fixture.Path, CancellationToken.None);

        Assert.Equal(installedVersion, result.CurrentVersion);
        Assert.Equal("1.2.4", result.LatestVersion);
        Assert.Equal("up-to-date", result.Status);
        Assert.Null(result.LogPath);
        Assert.Equal(0, handoffCount);
        Assert.Equal(0, confirmationCount);
        Assert.Single(handler.Requests);
    }

    [Fact]
    public async Task Tampered_payload_is_rejected_without_handoff()
    {
        using var fixture = new OwnedDirectory();
        var expectedPayload = Encoding.UTF8.GetBytes("expected installer");
        var tamperedPayload = Encoding.UTF8.GetBytes("tampered installer");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            if (request.RequestUri!.AbsoluteUri == LatestUrl)
            {
                return Reply(JsonResponse(ReleaseJson(
                    "1.2.4", new object[] { Asset(123, installerName, Sha256Digest(expectedPayload)) })));
            }

            if (request.RequestUri.AbsoluteUri == AssetUrl)
            {
                AssertApiRequest(request, "application/octet-stream");
                return Reply(BinaryResponse(tamperedPayload));
            }

            return Failure<HttpResponseMessage>("Unexpected request: " + request.RequestUri);
        });
        using var client = new HttpClient(handler);
        var handoffCount = 0;
        var updater = new ReleaseUpdater(client, (_, _) =>
        {
            handoffCount++;
            return "must-not-start";
        }, AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Equal(0, handoffCount);
        Assert.Equal([LatestUrl, AssetUrl], handler.Requests.Select(request => request.Uri.AbsoluteUri).ToArray());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("sha1:0000000000000000000000000000000000000000000000000000000000000000")]
    [InlineData("sha256:")]
    [InlineData("sha256:000000000000000000000000000000000000000000000000000000000000000")]
    [InlineData("sha256:00000000000000000000000000000000000000000000000000000000000000000g")]
    public async Task Missing_or_wrong_digest_is_rejected_before_download(string? digest)
    {
        using var fixture = new OwnedDirectory();
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(ReleaseJson(
                "1.2.4", new object[] { Asset(123, installerName, digest) })));
        });
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Single(handler.Requests);
    }

    [Fact]
    public async Task Missing_exact_installer_asset_is_rejected()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(ReleaseJson(
                "1.2.4",
                new object[] { Asset(123, "outlook-cli-1.2.4-setup.exe.asc", Sha256Digest(payload)) })));
        });
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Single(handler.Requests);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public async Task Non_positive_installer_asset_id_is_rejected(long assetId)
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(ReleaseJson(
                "1.2.4", new object[] { Asset(assetId, installerName, Sha256Digest(payload)) })));
        });
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Single(handler.Requests);
    }

    [Fact]
    public async Task Duplicate_exact_installer_assets_are_rejected_as_invalid_metadata()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(ReleaseJson(
                "1.2.4",
                new object[]
                {
                    Asset(123, installerName, Sha256Digest(payload)),
                    Asset(124, installerName, Sha256Digest(payload)),
                })));
        });
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Single(handler.Requests);
    }

    [Theory]
    [InlineData("not-json")]
    [InlineData("{}")]
    [InlineData("{\"tag_name\":\"1.2.4\",\"assets\":[]}")]
    [InlineData("{\"tag_name\":\"v1.2\",\"assets\":[]}")]
    [InlineData("{\"tag_name\":\"v1.2.4.0\",\"assets\":[]}")]
    [InlineData("{\"tag_name\":\"V1.2.4\",\"assets\":[]}")]
    [InlineData("{\"tag_name\":\"v1.2.4\",\"draft\":true,\"assets\":[]}")]
    [InlineData("{\"tag_name\":\"v1.2.4\",\"prerelease\":true,\"assets\":[]}")]
    public async Task Invalid_release_metadata_is_reported_as_invalid_operation(string body)
    {
        using var fixture = new OwnedDirectory();
        var handler = new RecordingHandler(request =>
        {
            Assert.Equal(LatestUrl, request.RequestUri!.AbsoluteUri);
            return Reply(JsonResponse(body));
        });
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Single(handler.Requests);
    }

    [Fact]
    public async Task Http_failure_is_propagated_without_handoff()
    {
        using var fixture = new OwnedDirectory();
        var expected = new HttpRequestException("network failure");
        var handler = new RecordingHandler(_ => Failure<HttpResponseMessage>(expected));
        using var client = new HttpClient(handler);
        var handoffCount = 0;
        var updater = new ReleaseUpdater(client, (_, _) =>
        {
            handoffCount++;
            return "must-not-start";
        }, AllowUpdate);

        var actual = await Assert.ThrowsAsync<HttpRequestException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Same(expected, actual);
        Assert.Equal(0, handoffCount);
    }

    [Fact]
    public async Task Non_success_http_status_is_propagated_as_http_request_exception()
    {
        using var fixture = new OwnedDirectory();
        var handler = new RecordingHandler(_ => Reply(new HttpResponseMessage(HttpStatusCode.InternalServerError)));
        using var client = new HttpClient(handler);
        var updater = new ReleaseUpdater(client, (_, _) => "must-not-start", AllowUpdate);

        await Assert.ThrowsAsync<HttpRequestException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));
    }

    [Fact]
    public async Task Non_https_redirect_is_rejected_without_following_or_handoff()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var insecureUrl = "http://downloads.example.test/" + installerName;
        var handler = new RecordingHandler(request =>
        {
            if (request.RequestUri!.AbsoluteUri == LatestUrl)
            {
                return Reply(JsonResponse(ReleaseJson(
                    "1.2.4", new object[] { Asset(123, installerName, Sha256Digest(payload)) })));
            }

            if (request.RequestUri.AbsoluteUri == AssetUrl)
            {
                AssertApiRequest(request, "application/octet-stream");
                return Reply(RedirectResponse(insecureUrl));
            }

            return Failure<HttpResponseMessage>("The insecure redirect must not be followed");
        });
        using var client = new HttpClient(handler);
        var handoffCount = 0;
        var updater = new ReleaseUpdater(client, (_, _) =>
        {
            handoffCount++;
            return "must-not-start";
        }, AllowUpdate);

        await Assert.ThrowsAsync<InvalidOperationException>(() => updater.UpdateAsync(
            new Version(1, 2, 3), fixture.Path, CancellationToken.None));

        Assert.Equal(0, handoffCount);
        Assert.Equal([LatestUrl, AssetUrl], handler.Requests.Select(request => request.Uri.AbsoluteUri).ToArray());
    }

    [Fact]
    public async Task Each_download_uses_a_unique_staging_directory()
    {
        using var fixture = new OwnedDirectory();
        var payload = Encoding.UTF8.GetBytes("installer bytes");
        var installerName = "outlook-cli-1.2.4-setup.exe";
        var handler = new RecordingHandler(request =>
        {
            if (request.RequestUri!.AbsoluteUri == LatestUrl)
            {
                return Reply(JsonResponse(ReleaseJson(
                    "1.2.4", new object[] { Asset(123, installerName, Sha256Digest(payload)) })));
            }

            if (request.RequestUri.AbsoluteUri == AssetUrl)
            {
                return Reply(BinaryResponse(payload));
            }

            return Failure<HttpResponseMessage>("Unexpected request: " + request.RequestUri);
        });
        using var client = new HttpClient(handler);
        var stagingDirectories = new List<string>();
        var updater = new ReleaseUpdater(client, (path, _) =>
        {
            var directory = Directory.GetParent(path)!.FullName;
            stagingDirectories.Add(directory);
            return "log";
        }, AllowUpdate);

        await updater.UpdateAsync(new Version(1, 2, 3), fixture.Path, CancellationToken.None);
        await updater.UpdateAsync(new Version(1, 2, 3), fixture.Path, CancellationToken.None);

        Assert.Equal(2, stagingDirectories.Distinct(StringComparer.OrdinalIgnoreCase).Count());
    }

    private static bool AllowUpdate(string currentVersion, string latestVersion) => true;

    private static void AssertApiRequest(HttpRequestMessage request, string accept)
    {
        Assert.Equal("GET", request.Method.Method);
        Assert.Equal("https", request.RequestUri!.Scheme);
        Assert.Equal("api.github.com", request.RequestUri.Host);
        Assert.Null(request.Headers.Authorization);
        Assert.Contains(request.Headers.Accept, value => value.MediaType == accept);
    }

    private static object Asset(long id, string name, string? digest) => new
    {
        id,
        name,
        digest,
    };

    private static string ReleaseJson(string tag, object? assets, bool draft = false, bool prerelease = false) =>
        JsonSerializer.Serialize(new
        {
            tag_name = "v" + tag,
            draft,
            prerelease,
            assets,
        });

    private static string Sha256Digest(byte[] bytes)
    {
        using var sha256 = SHA256.Create();
        return "sha256:" + string.Concat(sha256.ComputeHash(bytes).Select(value => value.ToString("x2")));
    }

    private static HttpResponseMessage JsonResponse(string body) =>
        new(HttpStatusCode.OK)
        {
            Content = new StringContent(body, Encoding.UTF8, "application/json"),
        };

    private static HttpResponseMessage BinaryResponse(byte[] payload) =>
        new(HttpStatusCode.OK)
        {
            Content = new ByteArrayContent(payload),
        };

    private static HttpResponseMessage RedirectResponse(string location) =>
        new(HttpStatusCode.Redirect)
        {
            Headers = { Location = new Uri(location) },
        };

    private static Task<T> Reply<T>(T value) => Task.FromResult(value);

    private static Task<T> Failure<T>(string message) =>
        Task.FromException<T>(new InvalidOperationException(message));

    private static Task<T> Failure<T>(Exception exception) => Task.FromException<T>(exception);

    private sealed class RecordingHandler(
        Func<HttpRequestMessage, Task<HttpResponseMessage>> responder) : HttpMessageHandler
    {
        public List<ObservedRequest> Requests { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(
            HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Assert.Null(request.Headers.Authorization);
            Requests.Add(new ObservedRequest(
                request.Method,
                request.RequestUri!,
                request.Headers.Authorization?.ToString(),
                request.Headers.Accept.Select(value => value.MediaType).ToArray()));
            return responder(request);
        }
    }

    private sealed class ObservedRequest(
        HttpMethod method,
        Uri uri,
        string? authorization,
        string?[] accept)
    {
        public HttpMethod Method { get; } = method;

        public Uri Uri { get; } = uri;

        public string? Authorization { get; } = authorization;

        public string?[] Accept { get; } = accept;
    }

    private sealed class OwnedDirectory : IDisposable
    {
        public OwnedDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "outlook-cli-release-updater-tests-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
            {
                Directory.Delete(Path, recursive: true);
            }
        }
    }
}
