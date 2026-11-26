using System.Diagnostics;
using Microsoft.Win32;

namespace OutlookCli.Shim;

internal static class Program
{
    private const string RegistryPath = @"Software\OutlookCli\Shim";
    private const string RegistryValueName = "Target";

    private static int Main()
    {
        try
        {
            var target = ReadTargetPath();
            var shimPath = typeof(Program).Assembly.Location;
            if (string.Equals(Path.GetFullPath(target), shimPath, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The configured target is the shim itself.");
            }

            var startInfo = new ProcessStartInfo(target, GetRawArgumentsAfterExecutable())
            {
                UseShellExecute = false,
            };

            // Keep the shim in the same console process group so Ctrl+C is sent to
            // the target as well. The handler prevents the shim from exiting first.
            ConsoleCancelEventHandler cancelHandler = (_, eventArgs) => eventArgs.Cancel = true;
            Console.CancelKeyPress += cancelHandler;
            try
            {
                using var process = Process.Start(startInfo)
                    ?? throw new InvalidOperationException("The configured target could not be started.");
                process.WaitForExit();
                return process.ExitCode;
            }
            finally
            {
                Console.CancelKeyPress -= cancelHandler;
            }
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"outlook-cli-shim: {exception.Message}");
            return 1;
        }
    }

    private static string ReadTargetPath()
    {
        using var key = Registry.CurrentUser.OpenSubKey(RegistryPath);
        var value = key?.GetValue(RegistryValueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
        if (value is not string target || string.IsNullOrWhiteSpace(target))
        {
            throw new InvalidOperationException(
                $"Registry value HKCU\\{RegistryPath}\\{RegistryValueName} is missing or empty.");
        }

        if (!IsAbsoluteWindowsPath(target))
        {
            throw new InvalidOperationException("The configured target path must be absolute.");
        }

        return target;
    }

    private static bool IsAbsoluteWindowsPath(string path)
    {
        return path.StartsWith(@"\\", StringComparison.Ordinal)
            || (path.Length >= 3
                && char.IsLetter(path[0])
                && path[1] == ':'
                && (path[2] == '\\' || path[2] == '/'));
    }

    private static string GetRawArgumentsAfterExecutable()
    {
        // Leave all quoting after argv[0] intact for the target process.
        var commandLine = Environment.CommandLine.TrimStart(' ', '\t');
        var inQuotes = false;
        for (var index = 0; index < commandLine.Length; index++)
        {
            var character = commandLine[index];
            if (character == '"')
            {
                inQuotes = !inQuotes;
            }
            else if (!inQuotes && (character is ' ' or '\t'))
            {
                return commandLine.Substring(index);
            }
        }

        return string.Empty;
    }
}
