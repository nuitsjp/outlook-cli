using System.Text.Json;

namespace OutlookCli.Commands;

internal sealed class JsonOutput(TextWriter output)
{
    private static readonly JsonSerializerOptions Options = new(JsonSerializerDefaults.Web);

    public void WriteLine<T>(T value) => output.WriteLine(JsonSerializer.Serialize(value, Options));

    public void Flush() => output.Flush();
}
