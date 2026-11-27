using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace OutlookCli.Updates;

internal sealed class UpdateInstaller
{
    private const string InstallerPathVariable = "OUTLOOK_CLI_UPDATE_INSTALLER_PATH";
    private const string InstallDirectoryVariable = "OUTLOOK_CLI_UPDATE_INSTALL_DIRECTORY";
    private const string LogPathVariable = "OUTLOOK_CLI_UPDATE_LOG_PATH";
    private const string CliProcessIdVariable = "OUTLOOK_CLI_UPDATE_CLI_PID";
    private const string CliStartTicksVariable = "OUTLOOK_CLI_UPDATE_CLI_START_TICKS";
    private const string ShimProcessIdVariable = "OUTLOOK_CLI_UPDATE_SHIM_PID";
    private const string ShimStartTicksVariable = "OUTLOOK_CLI_UPDATE_SHIM_START_TICKS";

    private const string ShimProcessIdEnvironmentVariable = "OUTLOOK_CLI_SHIM_PID";
    private const string ShimStartTicksEnvironmentVariable = "OUTLOOK_CLI_SHIM_START_TICKS";

    // The worker has no input other than the environment variables below.  Keeping
    // the script fixed means that installer and directory paths never become part
    // of a PowerShell command line.
    private const string WorkerScript = """
$ErrorActionPreference = 'Stop'

$installerPath = [Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_INSTALLER_PATH')
$installDirectory = [Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_INSTALL_DIRECTORY')
$logPath = [Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_LOG_PATH')

function Write-UpdateLog([string] $message) {
    $line = '{0} {1}' -f [DateTime]::UtcNow.ToString('o'), $message
    Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Wait-ForMatchingProcess(
    [string] $description,
    [string] $processIdText,
    [string] $startTicksText) {
    if ([string]::IsNullOrWhiteSpace($processIdText)) {
        return
    }

    [int] $processId = 0
    [long] $expectedStartTicks = 0
    if (-not [int]::TryParse($processIdText, [ref] $processId) -or
        -not [long]::TryParse($startTicksText, [ref] $expectedStartTicks)) {
        Write-UpdateLog ('{0} process identity was invalid; it was not awaited.' -f $description)
        return
    }

    $process = $null
    try {
        # Retain the handle so the process cannot be confused with a reused PID.
        $process = [Diagnostics.Process]::GetProcessById($processId)
        $null = $process.Handle
        $actualStartTicks = $process.StartTime.ToUniversalTime().Ticks
        if ($actualStartTicks -ne $expectedStartTicks) {
            Write-UpdateLog ('{0} process identity no longer matched; it was not awaited.' -f $description)
            return
        }

        Write-UpdateLog ('Waiting for {0} process {1}.' -f $description, $processId)
        $process.WaitForExit()
        Write-UpdateLog ('{0} process {1} exited.' -f $description, $processId)
    }
    catch [ArgumentException] {
        Write-UpdateLog ('{0} process {1} was no longer active.' -f $description, $processId)
    }
    catch [InvalidOperationException] {
        Write-UpdateLog ('{0} process {1} could not be awaited.' -f $description, $processId)
    }
    finally {
        if ($null -ne $process) {
            $process.Dispose()
        }
    }
}

function Read-FileVersion([string] $path) {
    $version = [Diagnostics.FileVersionInfo]::GetVersionInfo($path)
    return [Version]::new($version.FileMajorPart, $version.FileMinorPart, $version.FileBuildPart)
}

$mutex = $null
$ownsMutex = $false
try {
    Write-UpdateLog 'Update worker started.'
    Wait-ForMatchingProcess 'CLI' `
        ([Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_CLI_PID')) `
        ([Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_CLI_START_TICKS'))
    Wait-ForMatchingProcess 'WinGet shim' `
        ([Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_SHIM_PID')) `
        ([Environment]::GetEnvironmentVariable('OUTLOOK_CLI_UPDATE_SHIM_START_TICKS'))

    $userSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $mutex = [Threading.Mutex]::new($false, ('Global\outlook-cli-update-{0}' -f $userSid))
    try {
        $ownsMutex = $mutex.WaitOne()
    }
    catch [Threading.AbandonedMutexException] {
        $ownsMutex = $true
    }
    $currentVersion = Read-FileVersion (Join-Path $installDirectory 'outlook-cli.exe')
    $latestVersion = Read-FileVersion $installerPath
    if ($currentVersion -ge $latestVersion) {
        Write-UpdateLog ('Installed version {0} already meets or exceeds {1}.' -f $currentVersion, $latestVersion)
        return
    }

    # A trailing backslash immediately before a closing quote would escape that
    # quote in the Windows command line.  Double only that trailing run.
    $directoryArgument = $installDirectory -replace '(\\+)$', '$1$1'
    $installerArguments = @(
        '/VERYSILENT',
        '/SUPPRESSMSGBOXES',
        '/NORESTART',
        '/SP-',
        '/NOCLOSEAPPLICATIONS',
        ('/DIR="{0}"' -f $directoryArgument)
    )
    $installerProcess = Start-Process -FilePath $installerPath `
        -ArgumentList $installerArguments -WindowStyle Hidden -Wait -PassThru
    Write-UpdateLog ('Installer exit code: {0}' -f $installerProcess.ExitCode)
}
catch {
    Write-UpdateLog ('Update worker error: {0}' -f $_.Exception.Message)
}
finally {
    if ($ownsMutex) {
        $mutex.ReleaseMutex()
    }
    if ($null -ne $mutex) {
        $mutex.Dispose()
    }
    try {
        if (Test-Path -LiteralPath $installerPath) {
            Remove-Item -LiteralPath $installerPath -ErrorAction Stop
            Write-UpdateLog 'Downloaded installer removed.'
        }
        else {
            Write-UpdateLog 'Downloaded installer was already absent.'
        }
        [IO.Directory]::Delete([IO.Path]::GetDirectoryName($installerPath))
    }
    catch {
        Write-UpdateLog ('Downloaded installer cleanup failed: {0}' -f $_.Exception.Message)
    }
}
""";

    public string Start(string installerPath, string installDirectory)
    {
        if (string.IsNullOrWhiteSpace(installerPath))
        {
            throw new ArgumentException("The installer path is required.", nameof(installerPath));
        }

        if (string.IsNullOrWhiteSpace(installDirectory))
        {
            throw new ArgumentException("The install directory is required.", nameof(installDirectory));
        }

        var absoluteInstallerPath = Path.GetFullPath(installerPath);
        var absoluteInstallDirectory = Path.GetFullPath(installDirectory);
        var logPath = Path.GetFullPath(Path.Combine(
            Path.GetTempPath(),
            "outlook-cli-update-" + Guid.NewGuid().ToString("N", CultureInfo.InvariantCulture) + ".log"));

        using var currentProcess = Process.GetCurrentProcess();
        var startInfo = new ProcessStartInfo
        {
            FileName = Path.Combine(
                Environment.SystemDirectory,
                "WindowsPowerShell",
                "v1.0",
                "powershell.exe"),
            Arguments = "-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " +
                Convert.ToBase64String(Encoding.Unicode.GetBytes(WorkerScript)),
            UseShellExecute = false,
            CreateNoWindow = true,
        };

        var environment = startInfo.EnvironmentVariables;
        environment[InstallerPathVariable] = absoluteInstallerPath;
        environment[InstallDirectoryVariable] = absoluteInstallDirectory;
        environment[LogPathVariable] = logPath;
        environment[CliProcessIdVariable] = currentProcess.Id.ToString(CultureInfo.InvariantCulture);
        environment[CliStartTicksVariable] = currentProcess.StartTime
            .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture);

        var shimProcessId = Environment.GetEnvironmentVariable(ShimProcessIdEnvironmentVariable);
        var shimStartTicks = Environment.GetEnvironmentVariable(ShimStartTicksEnvironmentVariable);
        if (!string.IsNullOrWhiteSpace(shimProcessId) && !string.IsNullOrWhiteSpace(shimStartTicks))
        {
            environment[ShimProcessIdVariable] = shimProcessId;
            environment[ShimStartTicksVariable] = shimStartTicks;
        }

        environment.Remove("GH_TOKEN");

        using var worker = Process.Start(startInfo)
            ?? throw new InvalidOperationException("The update worker could not be started.");
        return logPath;
    }

}
