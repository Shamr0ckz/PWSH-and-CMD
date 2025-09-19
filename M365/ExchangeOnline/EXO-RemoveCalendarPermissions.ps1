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
$CalendarUPN = Read-Host "Enter UPN of user's calendar to remove permissions"
$CalendarPath = $CalendarUPN + ":\Calendar"
Write-Host "Current permissions are:"
Get-MailboxFolderPermission -Identity $CalendarUPN
$CalendarDelegate = Read-Host "Enter UPN of user to remove sharing"

#Execute share#
Remove-MailboxFolderPermission -Identity $CalendarUPN -User $CalendarDelegate 
Write-Host "New permissions are:"
Get-MailboxFolderPermission -Identity $CalendarUPN
