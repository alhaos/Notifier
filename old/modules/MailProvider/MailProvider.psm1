using namespace MimeKit
using namespace MailKit.Net.Smtp

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$True}

Add-Type -Path $PSScriptRoot/MailKit.dll
Add-Type -Path $PSScriptRoot/MimeKit.dll

function Send-AccuMail {
    [CmdletBinding()]
    param(
        [string[]]$To,
        [string[]]$Cc,
        [string[]]$Bcc,
        [string]$Subject,
        [string]$HtmlBody
    )
    $SMTPClient = [SMTPClient]::new()
    $Message = [MimeMessage]::new()
    $BodyBuilder = [BodyBuilder]::new()
    $BodyBuilder.HtmlBody = $HtmlBody
    $Message.Body = $BodyBuilder.ToMessageBody()

    foreach ($toItem in $To) {
        $Message.To.Add($toItem)
    }

    foreach ($ccItem in $Cc) {
        $Message.Cc.Add($CcItem)
    }

    foreach ($bccItem in $Bcc) {
        $Message.Bcc.Add($BccItem)
    }

    $Message.Subject = $Subject
    
    $Message.From.Add("accu-note@accureference.com")

    try {
        $SMTPClient.Connect("webmail.accureference.com", 587, [MailKit.Security.SecureSocketOptions]::Auto)
        $SMTPClient.Authenticate("accu-note@ac.com", "widen-qmgBMw#")
        $SMTPClient.Send($Message)
    }
    catch {
        throw $_
    }
    finally {
        $SMTPClient.Disconnect($true)
    }
}