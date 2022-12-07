using module .\modules\Notifier\Notifier.psd1 

Import-Module .\modules\GitsLogger\GitsLogger.psd1 -Force
Import-Module .\modules\Mailer\Mailer.psd1 -force

$DebugPreference = 'Continue'
$ErrorActionPreference = "Stop"
$VerbosePreference = 'SilentlyContinue'
#$VerbosePreference = 'Continue'

$conf = Import-PowerShellDataFile .\conf.psd1

Write-LogInfo ("$($conf.Name) start")

$Notifier = [Notifier]::new($conf.Notifier)

$Notifier.LoadData()
$Notifier.FillReps()
$Notifier.FillClients()
$Notifier.SendReps()
$Notifier.SendClients()

Write-LogInfo ("$($conf.Name) finish")