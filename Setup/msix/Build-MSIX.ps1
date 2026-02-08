<#
.SYNOPSIS
Builds and packages GameBase.NET as an MSIX installer.

.DESCRIPTION
Builds Release binaries, generates MSIX assets from an ICO, creates the MSIX
package layout, and produces a signed MSIX package. Requires Windows SDK tools
(MakeAppx, SignTool) and Microsoft Access Database Engine for Access export.

.PARAMETER packageName
Package identity name.

.PARAMETER publisher
Publisher string for MSIX identity. Must match signing certificate subject.

.PARAMETER publisherDisplayName
Publisher display name shown in installer UI.

.PARAMETER displayName
Display name for the application.

.PARAMETER version
MSIX package version.

.PARAMETER appExecutable
Primary application executable (relative to staging root).

.PARAMETER iconPath
Path to the ICO used to generate MSIX assets.

.PARAMETER outputPath
Directory where the MSIX package will be written.

.PARAMETER configuration
Build configuration for dotnet build.

.PARAMETER sdkToolsPath
Optional path to Windows SDK tools (MakeAppx.exe, signtool.exe).

.PARAMETER createTestCertificate
Create a self-signed test certificate for signing.

.PARAMETER certificatePath
Path to an existing PFX certificate for signing.

.PARAMETER certificatePassword
Password for the PFX certificate.
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$packageName,

  [Parameter(Mandatory = $true)]
  [string]$publisher,

  [Parameter(Mandatory = $true)]
  [string]$publisherDisplayName,

  [Parameter(Mandatory = $true)]
  [string]$displayName,

  [Parameter(Mandatory = $true)]
  [string]$version,

  [Parameter(Mandatory = $true)]
  [string]$appExecutable,

  [Parameter(Mandatory = $true)]
  [string]$iconPath,

  [Parameter(Mandatory = $true)]
  [string]$outputPath,

  [Parameter(Mandatory = $false)]
  [ValidateSet('Release', 'Debug')]
  [string]$configuration = 'Release',

  [Parameter(Mandatory = $false)]
  [string]$sdkToolsPath,

  [Parameter(Mandatory = $false)]
  [switch]$createTestCertificate,

  [Parameter(Mandatory = $false)]
  [string]$certificatePath,

  [Parameter(Mandatory = $false)]
  [string]$certificatePassword
)

function Write-Log {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$message
  )
  Write-Output $message
}

function New-MsixAssets {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$sourceIconPath,

    [Parameter(Mandatory = $true)]
    [string]$assetsPath
  )

  if (-not (Test-Path -Path $assetsPath)) {
    New-Item -ItemType Directory -Path $assetsPath | Out-Null
  }

  Add-Type -AssemblyName System.Drawing

  $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($sourceIconPath)
  $bitmap = $icon.ToBitmap()

  $assetSpecs = @(
    @{ Name = 'Square150x150Logo.png'; Width = 150; Height = 150 },
    @{ Name = 'Square44x44Logo.png'; Width = 44; Height = 44 },
    @{ Name = 'Wide310x150Logo.png'; Width = 310; Height = 150 },
    @{ Name = 'StoreLogo.png'; Width = 50; Height = 50 }
  )

  foreach ($spec in $assetSpecs) {
    $destPath = Join-Path -Path $assetsPath -ChildPath $spec.Name
    $img = New-Object System.Drawing.Bitmap $spec.Width, $spec.Height
    $graphics = [System.Drawing.Graphics]::FromImage($img)
    $graphics.Clear([System.Drawing.Color]::Transparent)
    $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $graphics.DrawImage($bitmap, 0, 0, $spec.Width, $spec.Height)
    $img.Save($destPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $img.Dispose()
  }

  $bitmap.Dispose()
  $icon.Dispose()
}

function Get-ToolPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$toolName
  )

  if (-not [string]::IsNullOrWhiteSpace($sdkToolsPath)) {
    $candidate = Join-Path -Path $sdkToolsPath -ChildPath $toolName
    if (Test-Path -Path $candidate) {
      return $candidate
    }
  }

  $cmd = Get-Command -Name $toolName -ErrorAction SilentlyContinue
  if ($null -ne $cmd) {
    return $cmd.Source
  }

  throw "Required tool not found: $toolName. Install Windows SDK or provide -sdkToolsPath."
}

function Invoke-DotNetBuild {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$projectPaths,

    [Parameter(Mandatory = $true)]
    [string]$configurationName
  )

  foreach ($project in $projectPaths) {
    Write-Log "Building $project..."
    $args = @(
      'build',
      $project,
      '--configuration',
      $configurationName,
      '--nologo'
    )
    $result = & dotnet @args
    if ($LASTEXITCODE -ne 0) {
      throw "Build failed for $project"
    }
  }
}

function Copy-OutputFiles {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$stagingPath,

    [Parameter(Mandatory = $true)]
    [string[]]$projectOutputs
  )

  if (Test-Path -Path $stagingPath) {
    Remove-Item -Recurse -Force -Path $stagingPath
  }

  New-Item -ItemType Directory -Path $stagingPath | Out-Null

  foreach ($output in $projectOutputs) {
    if (-not (Test-Path -Path $output)) {
      throw "Output path not found: $output"
    }

    Get-ChildItem -Path $output -File | ForEach-Object {
      $dest = Join-Path -Path $stagingPath -ChildPath $_.Name
      Copy-Item -Path $_.FullName -Destination $dest -Force
    }
  }
}

function Get-ProjectOutputPath {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$projectRoot,

    [Parameter(Mandatory = $true)]
    [string]$configurationName
  )

  $basePath = Join-Path -Path $projectRoot -ChildPath ("bin\$configurationName")
  if (-not (Test-Path -Path $basePath)) {
    throw "Output base path not found: $basePath"
  }

  $tfmDir = Get-ChildItem -Path $basePath -Directory | Where-Object {
    $_.Name -like 'net8.0-windows*'
  } | Select-Object -First 1

  if ($null -eq $tfmDir) {
    throw "No net8.0-windows* output found under: $basePath"
  }

  return $tfmDir.FullName
}

function New-TestCertificate {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [string]$subject,

    [Parameter(Mandatory = $true)]
    [string]$pfxPath,

    [Parameter(Mandatory = $true)]
    [string]$pfxPassword
  )

  $cert = New-SelfSignedCertificate -Type Custom -Subject $subject -KeyUsage DigitalSignature -KeyAlgorithm RSA -KeyLength 2048 -CertStoreLocation 'Cert:\CurrentUser\My' -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3')
  $securePassword = ConvertTo-SecureString -String $pfxPassword -Force -AsPlainText
  Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $securePassword | Out-Null
}

try {
  $rootPath = Split-Path -Parent $PSScriptRoot
  $repoRoot = Split-Path -Parent $rootPath

  $manifestTemplate = Join-Path -Path $PSScriptRoot -ChildPath 'AppxManifest.xml'
  $stagingPath = Join-Path -Path $PSScriptRoot -ChildPath 'App'
  $assetsPath = Join-Path -Path $PSScriptRoot -ChildPath 'Assets'
  $outputDir = if ([System.IO.Path]::IsPathRooted($outputPath)) { $outputPath } else { Join-Path -Path $PSScriptRoot -ChildPath $outputPath }

  if (-not (Test-Path -Path $manifestTemplate)) {
    throw "Manifest template not found: $manifestTemplate"
  }

  if (-not (Test-Path -Path $iconPath)) {
    throw "Icon not found: $iconPath"
  }

  if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
  }

  $projectPaths = @(
    'src\GameBase\GameBase.csproj',
    'src\GBToolbox\GBToolbox.csproj',
    'src\GBLauncher\GBLauncher.csproj',
    'src\GBRegRun\GBRegRun.csproj',
    'src\HVSC2GB64\HVSC2GB64.csproj',
    'src\GBMigrator\GBMigrator.csproj',
    'src\BAC\BAC.csproj'
  ) | ForEach-Object { Join-Path -Path $repoRoot -ChildPath $_ }

  Invoke-DotNetBuild -projectPaths $projectPaths -configurationName $configuration

  $projectOutputs = @(
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\GameBase') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\GBToolbox') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\GBLauncher') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\GBRegRun') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\HVSC2GB64') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\GBMigrator') -configurationName $configuration),
    (Get-ProjectOutputPath -projectRoot (Join-Path -Path $repoRoot -ChildPath 'src\BAC') -configurationName $configuration)
  )

  Copy-OutputFiles -stagingPath $stagingPath -projectOutputs $projectOutputs

  New-MsixAssets -sourceIconPath $iconPath -assetsPath $assetsPath

  $manifestContent = Get-Content -Path $manifestTemplate -Raw
  $manifestContent = $manifestContent.Replace('__PACKAGE_NAME__', $packageName)
  $manifestContent = $manifestContent.Replace('__PUBLISHER__', $publisher)
  $manifestContent = $manifestContent.Replace('__PUBLISHER_DISPLAY__', $publisherDisplayName)
  $manifestContent = $manifestContent.Replace('__VERSION__', $version)
  $manifestContent = $manifestContent.Replace('__DISPLAY_NAME__', $displayName)
  $manifestContent = $manifestContent.Replace('__EXECUTABLE__', $appExecutable)

  $manifestPath = Join-Path -Path $stagingPath -ChildPath 'AppxManifest.xml'
  Set-Content -Path $manifestPath -Value $manifestContent -Encoding UTF8

  Copy-Item -Path $assetsPath -Destination (Join-Path -Path $stagingPath -ChildPath 'Assets') -Recurse -Force

  $makeAppx = Get-ToolPath -toolName 'MakeAppx.exe'
  $signTool = Get-ToolPath -toolName 'signtool.exe'

  $msixPath = Join-Path -Path $outputDir -ChildPath ("$packageName-$version.msix")

  $makeAppxArgs = @(
    'pack',
    '/d', $stagingPath,
    '/p', $msixPath,
    '/o'
  )
  & $makeAppx @makeAppxArgs
  if ($LASTEXITCODE -ne 0) {
    throw 'MakeAppx failed.'
  }

  if ($createTestCertificate -and [string]::IsNullOrWhiteSpace($certificatePath)) {
    $certificatePath = Join-Path -Path $outputDir -ChildPath 'GameBaseNet-TestCert.pfx'
    if ([string]::IsNullOrWhiteSpace($certificatePassword)) {
      $certificatePassword = Read-Host -Prompt 'Enter a password for the test certificate'
    }
    New-TestCertificate -subject $publisher -pfxPath $certificatePath -pfxPassword $certificatePassword
    Write-Log "Created test certificate at $certificatePath"
  }

  if (-not (Test-Path -Path $certificatePath)) {
    throw "Certificate not found: $certificatePath"
  }

  $signArgs = @(
    'sign',
    '/fd', 'SHA256',
    '/a',
    '/f', $certificatePath,
    '/p', $certificatePassword,
    $msixPath
  )
  & $signTool @signArgs
  if ($LASTEXITCODE -ne 0) {
    throw 'Signing failed.'
  }

  Write-Log "MSIX created: $msixPath"
}
catch {
  Write-Error $_
  throw
}
