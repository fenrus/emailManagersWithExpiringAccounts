# Script to email account holder that their password is expring within X days.
# 
#
# David Syk 2020-12-17
#

import-module ActiveDirectory;

#Set SMTP Server
$smtpServer = "smtp.server.here"
$senderEmail = "sender.details@yourdomain.com"


# where are the accounts located?
$searchroot = @("OU=Users,OU=SWEDEN,OU=Administrators,DC=domain,DC=local")
$searchroot += "OU=Users,OU=NORWAY,OU=Administrators,DC=domain,DC=local"


#logfile
$global:logfile = "C:\Scripts\emailPasswordReminder\log\emailPasswordReminder-$(get-date -f yyyyMMdd-HHmm).txt"

# Threshold to trigger warning - how many days prior to password expiration should a reminder email be sent
$dayswarning = "14"

$DebugPreference = "Continue"
$debugmail = "youremail@here.com"

function Write-Log
{
    param (
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('1','2','3')]
        [int]$Severity = 1 ## Default to a low severity. Otherwise, override
    )
    
    $message = "$(get-date -f yyyyMMdd-HHmm) $severity $message"

    #If debugging is turned on, write all the log output to screen. (as Debug/Warning/Error)
    if ($DebugPreference -eq "Continue") {
        if ($Severity -eq "1") {
            Write-Debug($message);
        }
        if ($Severity -eq "2") {
            Write-Warning($message);
        }
        if ($Severity -eq "3") {
            Write-Error($message);
        }    
    }

    # always log message to logfile.
    $message | Out-File -Append $logfile 
}

function send-infomail
{
    param (
        [Parameter(Mandatory)]
        [string]$Recipient,
        
        [Parameter(Mandatory)]
        [string]$samaccountname,

        [Parameter(Mandatory)]
        [string]$expirydate

        )

        $Body = "
                            <html>  
                            <body> 
                            <p>Hi <br/>
                            <br>You are being notified because our records show that you are the owner for the below listed <b>account</b>.<br>
                            <style>
                            TABLE {border-width: 1px; border-style: solid; border-color: black; border-=collapse: collapse;}
                            TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #FF9696;}
                            TD {border-width: 1px; padding: 1px; border-style: solid; border-color: black;}
                            </style>
                            <table>
                            <tbody>
                             <tr><th>samAccountName</th><th>PasswordExpires</th></tr>
                             <tr><td>$samAccountName</td><td>$expirydate</td></tr>
                             </tbody>
                             </table><br>

                             Please login and update your password prior to expiration to avoid login problems.<br />
                             This message will repeat weekly until the password is expired, or password reset.<br />
                             <br />
                             <br />
                             If you do not need your account, request deletion by contacting SD<br><br>
                                
                             <br />
                             <p>You can not reply to this email.</p>
                             
                             <p>Regards,<br />
                             Automation<br />

                             </body>  
                             </html>";
              
	# Create email object to send
	$msg = new-object Net.Mail.MailMessage
	$smtp = new-object Net.Mail.SmtpClient($smtpServer)
	$msg.From = $SenderEmail
	$msg.To.Add($recipient)
	#to see when going to production, BCC mail will be sent to debug mail.
        $msg.Bcc.Add($DebugMail)
	#email is in HTML format.
        $msg.IsBodyHTML = $true
        $msg.Subject = "Account: password expiration notification"
        #$msg.Priority = "High"
	# Set mail encoding to be UTF-8 to handle NO/DK/SE Chars.
	$msg.BodyEncoding = [System.Text.Encoding]::UTF8
	$msg.Body = $Body
	$smtp.Send($msg)
	Write-Log -Message "Reminder sent to $($Recipient) for $($samaccountname) expiring $($expirydate)" -Severity 2
}


# the loop over the OU's with accounts

for ($i=0; $i -lt $searchroot.length; $i++) {
       Write-Log -Message "Processing OU: [$($searchroot[$i])]"
       $sendEmail = $false
       
       Get-ADUser -Filter * -SearchBase $searchroot[$i] -Properties samaccountname,mail,msDS-UserPasswordExpiryTimeComputed | ForEach {
           $PwExpiryDate=[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")
           $PwExpiryDate1=$PwExpiryDate.ToShortDateString()
           $today=(Get-Date)
           $DaysLeft=($PwExpiryDate-$today).days
           Write-Log -Message "[$($_.samaccountname)] Has password expiry date: $($PwExpiryDate1)"

        if ($DaysLeft -ge 0 -and $DaysLeft -le $dayswarning) {
            Write-Log -Message "[$($_.samaccountname)] Password expiry happens in $($daysleft) days, which is less or equal to the threshold: $($dayswarning)."
            Write-Log -Message "[$($_.samaccountname)] Consider sending reminder"
            send-infomail -Recipient $_.mail -samaccountname $_.samaccountname -expirydate $PwExpiryDate1
        }
        else {
            Write-Log -Message "[$($_.samaccountname)] Password expiry happens in $($daysleft) days. If the value is negative it's number of days since the expiration"
        }
    }
} 

