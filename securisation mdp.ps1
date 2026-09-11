New-Item -ItemType Directory -Path "C:\mercator\Mercator" -Force
Get-Credential | Export-Clixml -Path "C:\Mercator\mercator_cred.xml"

