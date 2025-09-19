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

#Set and list current permissions#
$CalendarUPN = Read-Host "Enter UPN of user's calendar to remove permissions"
Write-Host "Current permissions are:"
Get-MailboxFolderPermission -Identity $CalendarUPN
$CalendarDelegate = Read-Host "Enter UPN of user to remove sharing"

#Execute share#
Remove-MailboxFolderPermission -Identity $CalendarUPN -User $CalendarDelegate 

#List new permissions#
Write-Host "New permissions are:"
Get-MailboxFolderPermission -Identity $CalendarUPN

Read-Host -Prompt "Press any key to end script..."
Exit
