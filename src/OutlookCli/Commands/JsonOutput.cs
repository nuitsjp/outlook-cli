using System.Text.Json;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

namespace OutlookCli.Commands;

internal sealed class JsonOutput(TextWriter output)
{
    private static readonly JsonSerializerOptions AsciiOptions = new(JsonSerializerDefaults.Web);
    private static readonly JsonSerializerSettings Utf8Settings = new()
    {
        ContractResolver = new CamelCasePropertyNamesContractResolver(),
    };

    public bool AsciiOnly { private get; set; }

    public void WriteLine<T>(T value)
    {
        // Keep the original ASCII format. Json.NET also writes supplementary Unicode characters literally.
        var line = AsciiOnly
            ? System.Text.Json.JsonSerializer.Serialize(value, AsciiOptions)
            : JsonConvert.SerializeObject(value, Utf8Settings);
        output.WriteLine(line);
    }

    public void Flush() => output.Flush();
}
