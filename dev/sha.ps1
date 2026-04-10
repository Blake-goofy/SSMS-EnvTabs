$vsix = "C:\Users\blake\source\repos\SSMS EnvTabs\SSMS EnvTabs\bin\Release\SSMS EnvTabs.vsix"
$hash = (Get-FileHash -Algorithm SHA256 $vsix).Hash.ToLower()
"$hash  $(Split-Path $vsix -Leaf)" | Set-Content -NoNewline "$vsix.sha256"