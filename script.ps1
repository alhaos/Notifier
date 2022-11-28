using assembly .\modules\DbProvider\System.Data.SQLite.dll
using module .\modules\LogProvider\LogProvider.psm1 
using module .\modules\DbProvider\DbProvider.psm1

$PSDefaultParameterValues."Import-Module:Force" = $true
$DebugPreference = 'Continue'
$ErrorActionPreference = "Stop"

$conf = Import-PowerShellDataFile .\conf.psd1
Write-LogInfo ("$($conf.Name) start")

$DbProvider = [DbProvider]::new($conf.Notifier.ConnectionString)
$DbProvider.GetClientIdArray()

Write-LogInfo ("$($conf.Name) end")
