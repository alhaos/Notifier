using assembly .\modules\DbProvider\System.Data.SQLite.dll
using module .\modules\LogProvider\LogProvider.psm1 
using module .\modules\DbProvider\DbProvider.psm1
using module .\modules\Report\Report.psm1

Import-Module .\modules\MailProvider\MailProvider.psm1 -Force

$PSDefaultParameterValues."Import-Module:Force" = $true
$DebugPreference = 'Continue'
$ErrorActionPreference = "Stop"

$conf = Import-PowerShellDataFile .\conf.psd1
Write-LogInfo ("$($conf.Name) start")

#$DbProvider = [DbProvider]::new($conf.Notifier.ConnectionString)
#$DbProvider.LoadData()

$Report = [Report]::new()
$Report.FillReps()
$Report.FillClients()

$Report.Clients.ForEach{

    $splat = @{
        To = @("alhaos@gmail.com")
        Cc = @("alhaos@gmail.com")
        Bcc = @("alhaos@gmail.com")
        Subject = "Test email"
        HtmlBody = $_.GetHtmlBody()
    }
    
    Send-AccuMail @splat
}

Write-LogInfo ("$($conf.Name) end")
