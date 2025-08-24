param (
  [string]$InstallPath = "$env:APPDATA\YourAddinName"
)

$ManifestPath = "$InstallPath\manifest.xml"
$SourcePath = "dist"

if (-not (Test-Path $InstallPath)) {
  New-Item -ItemType Directory -Path $InstallPath
}

Copy-Item -Path "$SourcePath\*" -Destination $InstallPath -Recurse

$RegistryPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$Name = "YourAddinName"
$Value = $ManifestPath

if (-not (Test-Path $RegistryPath)) {
  New-Item -Path $RegistryPath -Force
}

Set-ItemProperty -Path $RegistryPath -Name $Name -Value $Value
