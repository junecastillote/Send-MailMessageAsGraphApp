Function Send-MailMessageAsGraphApp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $MailFrom,

        [Parameter()]
        [string[]]
        $MailTo,

        [Parameter()]
        [string[]]
        $MailCc,

        [Parameter()]
        [string[]]
        $MailBcc,

        [Parameter(Mandatory)]
        [string]
        $MailSubject,

        [Parameter(Mandatory)]
        [string]
        $MailMessage,

        [Parameter()]
        [string[]]
        $MailAttachment
    )

    if (!($MailFrom)) {
        "The [MailFrom] address is required." | Write-Error
        return $null
    }

    if (!$MailTo -and !$MailCc -and !$MailBcc) {
        "At least one recipients [MailTo, MailCc, MailBcc] is required." | Write-Error
        return $null
    }

    Function ToEmailAddressHashTable {
        param(
            [parameter()]
            [string[]]
            $Address
        )

        $Address | ForEach-Object {
            @{
                EmailAddress = @{
                    Address = $_
                }
            }
        }
    }

    Function ToAttachmentHashTable {
        param (
            [Parameter()]
            [string[]]
            $Path
        )

        $Path | ForEach-Object {
            try {
                $filename = (Resolve-Path $_ -ErrorAction STOP).Path
            }
            catch {
                $_.Exception.Message | Write-Error
                return $null
            }

            if ($PSVersionTable.PSEdition -eq 'Core') {
                $file_b64_string = $([convert]::ToBase64String((Get-Content $filename -AsByteStream)))
            }
            else {
                $file_b64_string = $([convert]::ToBase64String((Get-Content $filename -Raw -Encoding byte)))
            }


            @{
                "@odata.type" = "#microsoft.graph.fileAttachment"
                Name          = (Split-Path $_ -Leaf)
                ContentBytes  = $file_b64_string
            }
        }
    }

    $mail_params = @{
        Message = @{
            Subject                = $MailSubject
            Body                   = @{
                ContentType = "HTML"
                Content     = $MailMessage
            }
            InternetMessageHeaders = @(
                @{
                    Name  = "X-Mailer"
                    Value = "SendMailMessageAsGraphApp"
                }
            )
        }
    }

    if ($MailTo) {
        $mail_params.Message.Add('toRecipients', @(ToEmailAddressHashTable $MailTo))
    }

    if ($MailCc) {
        $mail_params.Message.Add('ccRecipients', @(ToEmailAddressHashTable $MailCc))
    }

    if ($MailBcc) {
        $mail_params.Message.Add('bccRecipients', @(ToEmailAddressHashTable $MailBcc))
    }

    if ($MailAttachment) {
        $mail_params.Message.Add('Attachments', @(ToAttachmentHashTable $MailAttachment))
    }

    Send-MgUserMail @mail_params -UserId $MailFrom
}
