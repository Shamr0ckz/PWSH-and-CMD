#Set-ExecutionPolicy RemoteSigned -Force -Confirm:$false;

#Check for Exchannge Online module and install if missing#
if(-not (Get-Module ExchangeOnlineManagement -ListAvailable)){
Write-Host "Installing Exchange Online module for current user"
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
} else {
	write-host "Exchange Online module already installed"
}

#Connect to Exchange Online (EXO)#
$AdminUPN = Read-Host "Enter your Exchange Online admin UPN"
Connect-ExchangeOnline -UserPrincipalName $AdminUpn

#Set variables#
$CalendarUPN = Read-Host "Enter UPN of user's calendar to share"
$CalendarPath = $CalendarUPN + ":\Calendar"
$CalendarDelegate = Read-Host "Enter UPN of delegate to share the calendar with"
$Permissions = Read-Host "Enter permissions to grant" $CalendarDelegate "(Owner, Editor, Reviewer, etc.)"

#Execute share#
Add-MailboxFolderPermission -Identity $CalendarPath -User $CalendarDelegate -AccessRights $Permissions

#Show new permissions#
Write-Host "New permissions are:"
Get-MailboxFolderPermission -Identity $CalendarUPN

Read-Host -Prompt "Press any key to end script..."
Exit
