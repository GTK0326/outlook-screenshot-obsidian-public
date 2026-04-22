[CmdletBinding()]
param(
    [string]$PackageRoot,
    [switch]$SkipPathUpdate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-CSharpCompilerPath {
    $candidates = @(
        "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
        "C:\Windows\Microsoft.NET\Framework\v4.0.30319\csc.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    throw "csc.exe was not found. Install .NET Framework build tools or compile OutlookScreenshotLauncher.cs manually."
}

function Add-DirectoryToUserPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DirectoryPath
    )

    $currentPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::User)
    $segments = @()
    if (-not [string]::IsNullOrWhiteSpace($currentPath)) {
        $segments = $currentPath.Split(";") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    }

    $normalizedSegments = $segments | ForEach-Object { $_.TrimEnd("\") }
    $normalizedTarget = $DirectoryPath.TrimEnd("\")
    if ($normalizedSegments -contains $normalizedTarget) {
        return $false
    }

    $updatedSegments = @($segments + $DirectoryPath)
    [Environment]::SetEnvironmentVariable("Path", ($updatedSegments -join ";"), [EnvironmentVariableTarget]::User)
    return $true
}

if ([string]::IsNullOrWhiteSpace($PackageRoot)) {
    $PackageRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}

$resolvedPackageRoot = [IO.Path]::GetFullPath($PackageRoot)
$configSamplePath = Join-Path $resolvedPackageRoot "outlook-screenshot.sample.json"
$configPath = Join-Path $resolvedPackageRoot "outlook-screenshot.json"
$launcherSourcePath = Join-Path $resolvedPackageRoot "OutlookScreenshotLauncher.cs"
$launcherExePath = Join-Path $resolvedPackageRoot "outlook-screenshot.exe"

if (-not (Test-Path -LiteralPath $configSamplePath)) {
    throw "Sample config was not found: $configSamplePath"
}

if (-not (Test-Path -LiteralPath $launcherSourcePath)) {
    throw "Launcher source was not found: $launcherSourcePath"
}

if (-not (Test-Path -LiteralPath $configPath)) {
    Copy-Item -LiteralPath $configSamplePath -Destination $configPath
    Write-Host ("Created config: {0}" -f $configPath)
}

$compilerPath = Get-CSharpCompilerPath
& $compilerPath /nologo /target:winexe /out:$launcherExePath /r:System.Windows.Forms.dll $launcherSourcePath
Write-Host ("Built launcher: {0}" -f $launcherExePath)

if (-not $SkipPathUpdate) {
    $pathUpdated = Add-DirectoryToUserPath -DirectoryPath $resolvedPackageRoot
    if ($pathUpdated) {
        Write-Host ("Added to user PATH: {0}" -f $resolvedPackageRoot)
    }
    else {
        Write-Host ("Already in user PATH: {0}" -f $resolvedPackageRoot)
    }
}

Write-Host ""
Write-Host "Next steps:"
Write-Host ("1. Edit {0}" -f $configPath)
Write-Host "2. Set OPENAI_API_KEY"
Write-Host "3. Restart Explorer or sign out/in if PATH was updated"
Write-Host "4. Run: outlook-screenshot"
