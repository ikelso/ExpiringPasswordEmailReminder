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
$smtpServer="server.fqdn.com"
$from = "FromEmail@fqdn.com"
$logging = "Enabled" # Set to Disabled to Disable Logging
$logFile = "C:\logs\ExpiredPasswords.log" # ie. c:\mylog.csv
$testing = "Disabled" # Set to Disabled to Email Users
$testRecipient = "testrecipient@fqdn.com"
$date = Get-Date -format ddMMyyyy

### The following variables are used in the html formatted email to end users
$ServiceDeskPhoneNumber =  "(555) 555-5555"
$ServiceDeskExtension = "5555"
$ServiceDeskEmail = "servicedesk@fqdn.com"
$CompanyName = "Non-Descript MSP"
$DepartmentName = "Information Technology"
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
$users = get-aduser -filter {(Enabled -eq "True") -and (PasswordNeverExpires -eq "False") -and (PasswordExpired -eq "False")} -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress
$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

# Process Each User for Password Expiry
foreach ($user in $users)
{
	#Set $Name to a user's first name
    $Name = (Get-ADUser $user).GivenName
	#Set $emailaddress to a user's email address from AD
    $emailaddress = $user.emailaddress
	#get the date the user reset their password
    $passwordSetDate = (get-aduser $user -properties *).PasswordLastSet
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
		Your Password will expire $messageDaysBold Please follow the instructions below to change your password before it expires.  If you require assistance, please contact our Service Desk at internal only extension $ServiceDeskExtension or by dialing $ServiceDeskPhoneNumber.
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
		<strong>$CompanyName</strong>
		<br>
		<strong>$DepartmentName</strong>
		</p>

		<p>
		<strong>DIRECT:</strong>
		Internal only extension $ServiceDeskExtension or $ServiceDeskPhoneNumber
		<br>
		<strong>EMAIL:</strong>
		<a href=`"mailto:$ServiceDeskEmail`">$ServiceDeskEmail</a>
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
