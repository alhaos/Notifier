using namespace MailKit.Net.Smtp
using namespace MimeKit

Set-StrictMode -Version 'Latest'

[System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }

function Send-AccuMail {
    [CmdletBinding()]
    param (
        [string[]] $To,
        [string[]] $Cc,
        [string[]] $Bcc,
        [string] $HtmlBody,
        [string] $Subject

    )

    $SMTPClient = [SmtpClient]::New()

    $message = [Mimemessage]::new()
    $bodyBuilder = [BodyBuilder]::new()
    $bodyBuilder.HtmlBody = $HtmlBody

    $message.Body = $bodyBuilder.ToMessageBody()

    foreach ($toItem in $To) {
        $message.To.Add($toItem)
    }

    foreach ($ccItem in $Cc) {
        $message.Cc.Add($ccItem)
    }

    foreach ($bccItem in $Bcc) {
        $message.Bcc.Add($bccItem)
    }
    
    $message.Subject = $Subject
    
    $message.From.Add("accu-note@accureference.com")

    try {
        $SMTPClient.Connect("webmail.accureference.com", 587, [MailKit.Security.SecureSocketOptions]::Auto)
        $SMTPClient.Authenticate("accu-note@ac.com", "widen-qmgBMw#")
        $null = $SMTPClient.Send($message)
    }
    catch {
        Write-Error $_
        return $false
    }
    finally {
        $SMTPClient.Disconnect($true)
    }
    return $true
}
