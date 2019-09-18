# Script to email managers about their direct reporting accounts that are expiring.
# It produces one mail per manager with a list of accounts that expire within 40 days.
#
# The script loops over all managers and checks validity of their direct reports.
# Reply-to in the email is set to SD to handle responses.
#
# There is no check for ".ext" or "-ext" or "ext-" pre/suffix in the usernames.
# The check is only for expiring accounts with a configured manager!
#
# Found code here for inspiration: https://stackoverflow.com/questions/50971290/powershell-send-single-email-to-manager-for-all-his-soon-to-expire-users
# 
#
# 2019-09-17, David 'fenrus' Syk
# 
#

import-module ActiveDirectory;

#Set SMTP Server
$smtpServer = "your.smtp.server.here"
$senderEmail = "noreplyaddress@yourdomain.com"

#logfile
$global:logfile = "C:\Scripts\emailManagersWithExpiringAccounts\log\expiremail-$(get-date -f yyyyMMdd-HHmm).txt"

# Debugging preferences
#
# Set value to "Continue" for runs with console output and all emails to DebugMail
$DebugPreference = "Continue"
$DebugMail = "your@yourdomain.com"

# 
# Set to "SilentlyContinue" to send emails to managers. BE VERY CAREFUL.
#
#$DebugPreference = "SilentlyContinue"



## Enter the AD Searchroots in this array.
# We search multiple OU's for managers.
$searchroot = @("OU=Users1,OU=ONE,DC=your,DC=domain,DC=net")
$searchroot += "OU=Users2,OU=TWO,DC=your,DC=domain,DC=net"



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

# if we are debugging, let's set the email recipient address to our debug email.
if ($DebugPreference -eq "Continue") {
    $ToEmail = $DebugMail
    Write-Log -Message "Debugging is active, using [$($ToEmail)] "
}


# Main program loop
for ($i=0; $i -lt $searchroot.length; $i++) {
    Write-Log -Message "Processing OU: [$($searchroot[$i])]"
    #reset 
    $sendEmail = $false

    Get-ADUser -Filter * -SearchBase $searchroot[$i] -Properties directReports,EmailAddress,GivenName,Surname | ForEach {
        $ManagerName=$_.GivenName + " " + $_.Surname
        $ManagerAccount=$_.samaccountname
        $Body = "
                            <html>  
                            <body> 
                            <p>Hi $ManagerName<br/>
                            <br>You are being notified because our records show that you are the primary contact Manager for the below listed users.<br>
                            <style>
                            TABLE {border-width: 1px; border-style: solid; border-color: black; border-=collapse: collapse;}
                            TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
                            TD {border-width: 1px; padding: 1px; border-style: solid; border-color: black;}
                            </style>
                            <table>
                            <tbody>
                             <tr><th>Name</th><th>Email</th><th>Title</th><th>AccountExpires</th></tr>";
        $AddBody = "";

        If ( $_.directReports ) {
            Write-Log -Message "Processing manager: [$($ManagerAccount)]"

            # if we are running this in normal mode, "SilentlyContinue", Set the recipient to the manager email.
            if ($DebugPreference -eq "SilentlyContinue") {
                $ToEmail = $_.EmailAddress
                Write-Log -Message "Debugging is not active, using [$($ToEmail)] "
            }
           


    
            $_.directReports | ForEach {

                $userDetails = Get-ADUser $_ -Properties AccountExpirationDate,accountExpires,EmailAddress,Title
                $userAccount = $userDetails.SamAccountName

                $userName=$userdetails.Name
                $userEmail=$userdetails.userPrincipalName
                $Title=$userDetails.Title

                if ($DebugPreference -eq "Continue") {
                    Write-Log -Message "                 Processing [$($userAccount)] "
                }

                If( $userDetails.accountExpires -eq 0 -or $userDetails.accountExpires -eq 9223372036854775807 ) {
                    # The account is not expiring.

                    # Check if the sendEmail is already set true, if it is, dont overwrite it.
                    # important if you have both accounts with and without expirydate as direct reports on the same manager.
                    if ($sendEmail -ne $true) {  
                        $sendEmail = $false
                    }
                    if ($DebugPreference -eq "Continue") {
                        Write-Log -Message "                 Processing [$($userAccount)] does not expire"
                    }
                }

                If ( $userDetails.AccountExpirationDate ) {
                    # the account have some expiration date set.
                    $ExpiryDate=$userDetails.AccountExpirationDate
                    $ExpiryDate1=$ExpiryDate.ToShortDateString()

                    $today=(Get-Date)

                    $DaysLeft=($ExpiryDate-$today).days

                    if ($DebugPreference -eq "Continue") {
                        Write-Log -Message "                      Processing [$($userAccount)] expirydate is $($ExpiryDate1)"
                    }

                    If ($DaysLeft -le 40 -and $DaysLeft -ge 0) {
                        if ($DebugPreference -eq "Continue") {
                            Write-Log -Message "                      Processing [$($userAccount)] expirydate is less than, or equal to 40, and more than 0."
                        }
                        # if less than 40 days, let's add it to the list of accounts in the email.
                        $AddBody += "<tr><td>$userName</td> <td><a style='text-decoration:none;color: rgb(0, 0, 0);'>$userEmail</a></td><td>$Title</td><td>$ExpiryDate1</td> </tr>";
                        $sendEmail = $true
                    }
                }
            }

            If ( $sendEmail ) {
                # prepare the rest of the mail body.
                $Body +=$AddBody;
                $Body = $Body + "</tbody>
                                </table><br>

                                 Take action as soon as possible to extend the validity before expiration.<br />
                                Reply to this email with information if you would like to extend the account.<br><br>
                                If you want to extend the user account, the default and maximum extension is 1 year. <br>
                                If you want the account to be deleted, request that according to standard procedure.<br><br>
                                <a style='text-decoration:bold;color: rgb(255, 0, 0);'>If nothing is done the account will expire and stop working the above mentioned date.</a><br>
                                </p>
                                <p>Regards<br />
                                Service Desk<br>

                                </body>  
                                </html>";
              
                    # Create email object to send

                    $msg = new-object Net.Mail.MailMessage

                    $smtp = new-object Net.Mail.SmtpClient($smtpServer)

                    $msg.From = $SenderEmail

                    $msg.To.Add($ToEmail)

                    #to see when going to production, BCC mail will be sent to debug mail.
                    $msg.Bcc.Add($DebugMail)
                    
                    #email is in HTML format.
                    $msg.IsBodyHTML = $true
                    $msg.Subject = "Account expiry notification. Action is required"
                    $msg.Priority = "High"

                    #set the reply to all responses go to servicedesk for handling
                    $msg.ReplyTo = "servicedesk@yourdomain.com"
                
                    # Set mail encoding to be UTF-8 to handle special chars.
                    $msg.BodyEncoding = [System.Text.Encoding]::UTF8
                    $msg.Body = $Body

                    $smtp.Send($msg)
                    Write-Log -Severity 2 -Message "Emailed manager: [$($ManagerName)] on email [$($ToEmail)]"
                    
                    #clean-up
                    $sendEmail = $false
                }
        }
    }
}
