<#----------------------------------------------------
   Created by: Nicholas Glantz
   Date: 08/08/18
   Summary: This script automates the process of creating a new user in an Exchange domain. It will create the 
   Exchange mailbox first then modify the AD user to the appropriate uniformed fields according to site, title, department etc...
   ADMINISTRATOR privelages are required for this process.Please ensure that the Exchange server has the proper permission to allow remote command execution.
----------------------------------------------------#>
Import-Module ActiveDirectory 
$ScriptPath = Split-Path -parent $PSCommandPath


<#Asks for Admin credentials to make changes in Exchange#>
$ExchangeCredential = Get-Credential -Credential ServerAdmin


<# If selected no to importing Email this will prompt all the neccesary variables#>
$FirstName = Read-Host 'First Name'
$LastName = Read-Host 'Last Name'
$EmployeeID = Read-Host 'Employee Number'
$Site = Read-Host 'Site'
$JobTitle = Read-Host 'Job Title'
$Department = Read-Host 'Department'
 



$Logon = "$FirstName.$LastName"
$Email = "$FirstName.$LastName@contoso.net"


<#Sets the Drivepath and logon scripts to N/A by default#>
$LogonScript = "N/A"
$ProfilePath = "N/A"
$RemotePath = "N/A"
$DrivePath = "N/A"

<# Takes the first initaial of the user and employee id and creates a password#>
$FirstInit = ($FirstName.tolower())[0]
$password = "$FirstInit$EmployeeID"
<#Converts the password into a SecureString#>
$securePass = ConvertTo-SecureString $password -AsPlainText -Force

<# Sets a Variable with the date and my initials. This variable will be used to fill the notes on the user object#>
$Notes = Get-Date -UFormat "%m/%d/%Y: NG"

<#Create User in Exchange#>
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://EXCHANGESERVER.contoso.net/PowerShell/ -Authentication Kerberos -Credential $ExchangeCredential
Import-PSSession $Session -WarningAction SilentlyContinue 
<# You can set which mailbox database the user will be in with -database parameter, the Address Book Policy -AddressBookPolicy, and Retention Policy -RetentionPolicy#>
new-mailbox -Alias "$FirstName.$Lastname" -FirstName "$FirstName" -LastName "$Lastname" -displayname "$FirstName $Lastname" -Name "$FirstName $Lastname"  -UserPrincipalName "$FirstName.$Lastname@ocdc.net" -Password $securePass
Remove-PSSession $Session

Start-Sleep -Seconds 5


    $LogonScript = "LOGONSCRIPT.bat"
    $ProfilePath = "\\Users\$FirstName.$LastName\Profile"
    $RemotePath = "\\Users\$FirstName.$LastName\Citrix"
    $DrivePath = "\\Users\$FirstName.$LastName"
    $OU = "OU=Users,DC=contoso,DC=net"
    $Description = "$JobTitle"
    $Company = "Contoso"

    <# Gives user group membership to county staff groups #>
    $User = Get-ADUser -Identity "$FirstName.$LastName" -Server "10COMMOSVR01.ocdc.net"
    $Group = Get-ADGroup -Identity "CN=*Marion and Clackamas County Staff,OU=Groups,OU=Marion and Clackamas County,OU=Counties,DC=ocdc,DC=net" -Server "10COMMOSVR01.ocdc.net"
    Add-ADGroupMember -Identity $Group -Members $User -Server "10COMMOSVR01.ocdc.net"
    $Group = Get-ADGroup -Identity "CN=Marion and Clackamas County Staff,OU=Groups,OU=Marion and Clackamas County,OU=Counties,DC=ocdc,DC=net" -Server "10COMMOSVR01.ocdc.net"
    Add-ADGroupMember -Identity $Group -Members $User -Server "10COMMOSVR01.ocdc.net"

<# Creates a Email template with the log on information#>
$CredentialTemplate = "
Email: $Email
Password: $password

"

"$CredentialTemplate" > newuser.txt

<#Open Notepad of Email Template#>
Invoke-Item newuser.txt

<# Sets Description, Profile Path - Home Drive - Drive Path, Company - Title#>
Set-ADUser -Identity "$FirstName.$LastName" -DisplayName "$FirstName $LastName ($Company)" -Description "$Description" -SamAccountName "$FirstName.$LastName" -UserPrincipalName "$FirstName.$LastName@ocdc.net" -EmailAddress "$Email" -Server "EXCHANGESERVER.contoso.net" -Enabled $true -Company "$Company" -ProfilePath "$ProfilePath" -ScriptPath "$LogonScript" -HomeDrive H -HomeDirectory "$DrivePath" -Title "$JobTitle" -GivenName "$FirstName" -Surname "$LastName" -Department "$Department"

<# Sets the IP phone and notes field is the user object#>
$i = Get-ADUser "$FirstName.$LastName" -Server "EXCHANGESERVER.contoso.net" -Properties info | %{ $_.info}
Set-ADUser "$FirstName.$LastName" -Server "EXCHANGESERVER.contoso.net" -Replace @{info="$($i) `r `n $Notes"}
Set-ADUser "$FirstName.$LastName" -Server "EXCHANGESERVER.contoso.net" -Replace @{ipPhone="$EmployeeID"}

Start-Sleep -Seconds 5

<#Sets Remote Services Profile Path #>
$user = [adsi]"LDAP://CN=$FirstName $LastName,$OU"
$User.psbase.invokeset("TerminalServicesProfilePath","$RemotePath")
$User.psbase.invokeset("TerminalServicesHomeDirectory","$DrivePath")
$user.psbase.invokeSet("TerminalServicesHomeDrive", "H:")

<#Disables "require user's permission" option in Remote Control properties#>
$user.psbase.invokeSet("EnableRemoteControl","2")

$user.setinfo()