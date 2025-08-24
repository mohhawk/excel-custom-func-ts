param (
  [string]$InstallPath = "$env:APPDATA\YourAddinName"
)

$RegistryPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$Name = "YourAddinName"

if (Test-Path $RegistryPath) {
  Remove-ItemProperty -Path $RegistryPath -Name $Name -ErrorAction SilentlyContinue
}

if (Test-Path $InstallPath) {
  Remove-Item $InstallPath -Recurse
}
