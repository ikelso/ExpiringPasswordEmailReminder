#################################################################################################################
#
# Version 1.0 April 2015
# Ian Kelso
# YetAnotherWindowsBlog.com
# Script to Automate Email Reminders when Users Passwords are due to Expire.
#
# Requires: Windows PowerShell Module for Active Directory
#
# For assistance and ideas, visit the TechNet Gallery Q&A Page. http://gallery.technet.microsoft.com/Password-Expiry-Email-177c3e27/view/Discussions#content
#
##################################################################################################################
# Please Configure the following variables....
$smtpServer="aus-dc-exch-01.usacompression.local"
$from = "IT Support itsupport@usacompression.com"
$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "C:\logs\ExpiredPasswords.log" # ie. c:\mylog.csv
$testing = "Disabled" # Set to Disabled to Email Users
$testRecipient = "ikelso@usacompression.com"
$date = Get-Date -format ddMMyyyy
#
###################################################################################################################



# Check Logging Settings
if (($logging) -eq "Enabled")
{
    # Test Log File Path
    $logfilePath = (Test-Path $logFile)
    if (($logFilePath) -ne "True")
    {
        # Create CSV File and Headers
        New-Item $logfile -ItemType File
        Add-Content $logfile "Date,Name,EmailAddress,DaystoExpire,ExpiresOn"
    }
} # End Logging Check

# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired
Import-Module ActiveDirectory
$users = get-aduser -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false }
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Process Each User for Password Expiry
foreach ($user in $users)
{
	#Set $Name to a user's first name
    $Name = (Get-ADUser $user).GivenName
	#Set $emailaddress to a user's email address from AD
    $emailaddress = $user.emailaddress
	#get the date the user reset their password
    $passwordSetDate = (get-aduser $user -properties * | foreach { $_.PasswordLastSet })
    $PasswordPol = (Get-AduserResultantPasswordPolicy $user)
    # Check for Fine Grained Password
    if (($PasswordPol) -ne $null)
    {
        $maxPasswordAge = ($PasswordPol).MaxPasswordAge
    }
 
    $expireson = $passwordsetdate + $maxPasswordAge
    $today = (get-date)
    $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
        
    # Set Greeting based on Number of Days to Expiry.

    # Check Number of Days to Expiry
    $messageDays = $daystoexpire

    if (($messageDays) -ge "1")
    {
        $messageDays = "in " + "$daystoexpire" + " days."
		$messageDaysBold = "in " + "<strong><font size=`"5`" color=`"red`">" + "$daystoexpire" + "</strong></font>" + " days."
    }
    else
    {
        $messageDays = "today."
		$messageDaysBold = "<strong><font size=`"5`" color=`"red`">" + "today." + "</strong></font>"
    }

    # Email Subject Set Here
    $subject="Your password will expire $messageDays"
 
    # Email Body Set Here, Note You can use HTML, including Images.
    $body ="
		<p>
		Dear $Name,
		</p>
		<p>
		Your Password will expire $messageDaysBold Please follow the instructions below to change your password before it expires.  If you require assistance, please contact our Service Desk at internal only extension 2001 or by dialing (877) 293-6378.
		</p>
		<p>
		<strong>Remember:</strong>
		</p>
		<ul>
		<li>
		Passwords must be at least 7 characters long
		</li>
		<li>
        You cannot reuse your last 6 passwords
		</li>
		<li>
		You must use 3 of the following: lower-case letters, upper-case letters, numbers, special characters
		</li>
		</ul>

		<p>
		<Strong>If you are in a networked office:</Strong>
		</p>
		<ul>
		<li>
		Press CTRL ALT Delete and select `"Change a password`.`.`.`" Enter your current password and your new password twice as instructed.
		</li>
		</ul>
		<p>
		<Strong> If you are outside an office: </Strong>
		</p>
		<ul>
		<li>
		Connect to the network via the VPN client (if it's before your expiration, you can ignore any prompts to change your password).
		</li>
		<li>
		Once connected, Press CTRL ALT Delete and select `"Change a password`.`.`.`" Enter your current password and your new password twice as instructed.
		</li>
		</ul>


		<p>
		<strong>USA Compression</strong>
		<br>
		<strong>Information Technology</strong>
		</p>

		<p>
		<strong>DIRECT:</strong>
		Internal only extension 2001 or (877) 293-6378
		<br>
		<strong>EMAIL:</strong>
		<a href=`"mailto:itsupport@usacompression.com`">itsupport@usacompression.com</a>
		</p>
		"

   
    # If Testing Is Enabled - Email Administrator
    if (($testing) -eq "Enabled")
    {
        $emailaddress = $testRecipient
    } # End Testing

    # If a user has no email address listed
    if (($emailaddress) -eq $null)
    {
        $emailaddress = $testRecipient    
    }# End No Valid Email

    # Send Email Message
    if ($daystoexpire -eq 7) 
    {
         # If Logging is Enabled Log Details
        if (($logging) -eq "Enabled")
        {
            Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson"
        }
        # Send Email Message
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High  

    } # End Send Message
	elseif ($daystoexpire -eq 4)
	{
         # If Logging is Enabled Log Details
        if (($logging) -eq "Enabled")
        {
            Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson"
        }
        # Send Email Message
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High  

    } # End Send Message
	elseif ($daystoexpire -eq 1)
    {
         # If Logging is Enabled Log Details
        if (($logging) -eq "Enabled")
        {
            Add-Content $logfile "$date,$Name,$emailaddress,$daystoExpire,$expireson"
        }
        # Send Email Message
        Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High  

    } # End Send Message
} # End User Processing



# End
# SIG # Begin signature block
# MIIIqgYJKoZIhvcNAQcCoIIImzCCCJcCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUakSnGH21F+7484YM0ZE7GDKg
# AbGgggX8MIIF+DCCBOCgAwIBAgITIQAAACdUWBVm5yWc2AAAAAAAJzANBgkqhkiG
# 9w0BAQUFADBgMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxHjAcBgoJkiaJk/IsZAEZ
# Fg5VU0FDb21wcmVzc2lvbjEnMCUGA1UEAxMeVVNBQ29tcHJlc3Npb24tQVVTLURD
# LUNBLTAxLUNBMB4XDTE1MDUwMzE1MTUxM1oXDTE2MDUwMjE1MTUxM1owdjEVMBMG
# CgmSJomT8ixkARkWBWxvY2FsMR4wHAYKCZImiZPyLGQBGRYOVVNBQ29tcHJlc3Np
# b24xITAfBgNVBAMTGE1hbmFnZWQgU2VydmljZSBBY2NvdW50czEaMBgGA1UEAxMR
# SWFuIEtlbHNvIChBZG1pbikwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC6Q9auqOjvSDTu3JX5fMCMeSlm6gWtMJmrFmCegsGlUsH3gTeCZuab3WDQ3Pwk
# 1UehngpyrCfSGw/g7rdhxyKjM24x1XozcYG4uKz7e+p67o2uA4dUl9x/Zn1VWjzO
# 6VB0MHtmf6ALOsYf3hW3/6RRHp//nFTvKBSwTO2TPpSfpE5N8M6m3PhCpR5FsaCp
# p5Cq7XFuo+KlNZr775O0UWbH3etookcyk8dzkHJYRtrZNnRAKXBDJgMB+WE/GKwT
# oLsZ/Oj6KasLSinAIN/mkOg2lGikS366tGJsMA58FlT7enZ+nvsaxSNeW16EkLHR
# r2XJgQXfdhJ/fGOsPKTvRcNXAgMBAAGjggKTMIICjzAlBgkrBgEEAYI3FAIEGB4W
# AEMAbwBkAGUAUwBpAGcAbgBpAG4AZzATBgNVHSUEDDAKBggrBgEFBQcDAzAOBgNV
# HQ8BAf8EBAMCB4AwHQYDVR0OBBYEFH+qJ4fD762Wo1IczROFXsRJ/vc3MB8GA1Ud
# IwQYMBaAFJiUE6Gx2meGHSZBnHyc18Sxge0KMIHqBgNVHR8EgeIwgd8wgdyggdmg
# gdaGgdNsZGFwOi8vL0NOPVVTQUNvbXByZXNzaW9uLUFVUy1EQy1DQS0wMS1DQSxD
# Tj1BVVMtREMtQ0EtMDEsQ049Q0RQLENOPVB1YmxpYyUyMEtleSUyMFNlcnZpY2Vz
# LENOPVNlcnZpY2VzLENOPUNvbmZpZ3VyYXRpb24sREM9VVNBQ29tcHJlc3Npb24s
# REM9bG9jYWw/Y2VydGlmaWNhdGVSZXZvY2F0aW9uTGlzdD9iYXNlP29iamVjdENs
# YXNzPWNSTERpc3RyaWJ1dGlvblBvaW50MIHZBggrBgEFBQcBAQSBzDCByTCBxgYI
# KwYBBQUHMAKGgblsZGFwOi8vL0NOPVVTQUNvbXByZXNzaW9uLUFVUy1EQy1DQS0w
# MS1DQSxDTj1BSUEsQ049UHVibGljJTIwS2V5JTIwU2VydmljZXMsQ049U2Vydmlj
# ZXMsQ049Q29uZmlndXJhdGlvbixEQz1VU0FDb21wcmVzc2lvbixEQz1sb2NhbD9j
# QUNlcnRpZmljYXRlP2Jhc2U/b2JqZWN0Q2xhc3M9Y2VydGlmaWNhdGlvbkF1dGhv
# cml0eTA4BgNVHREEMTAvoC0GCisGAQQBgjcUAgOgHwwdYS1pa2Vsc29AVVNBQ29t
# cHJlc3Npb24ubG9jYWwwDQYJKoZIhvcNAQEFBQADggEBAI2irthxT3GG9phRTqbA
# nDf23TXEZFHrCx24kcr63cj/1kax9pIXRihKNfzhcwzSc+7VpM9afGYzH0rIBuH5
# H7ny8eBigUN+gYlItXiG0cD/cdMnUNgWIH7vsJjjH8qzbh2TOXYDk2UxWai0NTlL
# c+rxTYbnto93i0fNDVq7ELrxi0uxIUQqFX2F3YLlY6MxyjL4KFfpAFMaooUijO8c
# ppTFLyE+gu1s4K2zFdnRaXjIHLuvG9uVFf1Le6+9OU2AWwAKPeJM1KycTTVotejC
# tfzYUwISE0Zs+v8c5UZBtB/kfYfIHrQFxGG/rzKPYhlswUG6bYxFJPeGZFpsPiOI
# WwwxggIYMIICFAIBATB3MGAxFTATBgoJkiaJk/IsZAEZFgVsb2NhbDEeMBwGCgmS
# JomT8ixkARkWDlVTQUNvbXByZXNzaW9uMScwJQYDVQQDEx5VU0FDb21wcmVzc2lv
# bi1BVVMtREMtQ0EtMDEtQ0ECEyEAAAAnVFgVZuclnNgAAAAAACcwCQYFKw4DAhoF
# AKB4MBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisG
# AQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcN
# AQkEMRYEFH2zmj+KuptkOTK8PXqzMZL5mXtbMA0GCSqGSIb3DQEBAQUABIIBACzZ
# F0BPBQoSGVSVgQFKEVrA2KLoKMzhesqR5Dl8+z2GbszFTy2T1f0UF9L1CSBtQfm0
# ioY/34cY+u/hdzpqGBkDZ3XiVQr2gZND6iOqkDRPWB9KCFMEb3tOoSUqr47CL4BN
# PtrpyB2gFeXODokFAixGDPgdBME4MXjFNGvu+gq8StlP/O8yIWqQF86Mc/QQeBg7
# SAhbsIAx9nzf8nW4Hiaz5igP2hELe5FDh2PCWjkferXOzhaOkvOn7CKWsNRUdWwI
# X+AQPB2s4XncbGEk/nSHerukVBPzUY920XJU8rQjFuBSsKj6etOQBnq5FHoJvgra
# tXuj2jrCSCrV1r73fBA=
# SIG # End signature block
