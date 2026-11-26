[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $Version
)

$ErrorActionPreference = 'Stop'
$repository = Split-Path -Parent $PSScriptRoot
$versionPattern = '(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)'

$remoteRefs = @(git -C $repository ls-remote --tags --refs --sort=-version:refname origin 'refs/tags/v*')
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

$stableTags = @($remoteRefs | ForEach-Object {
    if ($_ -cmatch "\trefs/tags/(v$versionPattern)$") {
        $Matches[1]
    }
})

if (-not $PSBoundParameters.ContainsKey('Version')) {
    $Version = '0.1.0'
    if ($stableTags.Count -gt 0) {
        $latest = [version]::Parse($stableTags[0].Substring(1))
        $Version = "$($latest.Major).$($latest.Minor).$($latest.Build + 1)"
    }
}

if ($Version -cnotmatch "^$versionPattern$" -or
    @($Version.Split('.') | Where-Object { [decimal]$_ -gt 65534 }).Count -gt 0) {
    throw 'Specify a version as X.Y.Z with each component between 0 and 65534.'
}

$tag = "v$Version"
if ($stableTags -ccontains $tag) {
    throw "Tag $tag already exists on origin."
}

git -C $repository tag $tag HEAD
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

git -C $repository push origin "refs/tags/$tag"
exit $LASTEXITCODE
