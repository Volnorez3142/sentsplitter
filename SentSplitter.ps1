#THIS SCRIPT ELEVATES THE STARTED PS SESSION TO ADMIN
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

#DECLARING THE DESKTOP VARIABLE PATH, STARTING THE TRANSCRIPT, SETTING MAILBOXES SENT COPY INTO BOTH ACCOUNTS TO $TRUE
$DesktopPath = [Environment]::GetFolderPath("Desktop")
Start-Transcript -path $DesktopPath\ExchangeConf.txt -append

#INSTALLING ALL THE DEPENDENCIES AND ALLOWING POLICIES
$installmoduleprompt = @"
**********************************************
* Do you wish to install PowershellGet,      *
* PSResourceGet and ExchangeOnlineManagement *
* modules?                                   *
*                  [yes/no]                  *
**********************************************
"@
Write-Host $installmoduleprompt -ForegroundColor Cyan
$installmoduleprompt = Read-Host

if ($installmoduleprompt -eq "yes") {
    Write-Host "Installing PowerShellGet..."
    Install-Module PowerShellGet -Force -AllowClobber
    Write-Host "Installing Powershell.PSResourceGet..."
    Install-Module Microsoft.PowerShell.PSResourceGet -Repository PSGallery
    Write-Host "Installing ExchangeOnlineManagement..."
    Install-Module -Name ExchangeOnlineManagement
    Write-Host "Setting Execution Policy to Unrestricted..."
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted
} else {
    Write-Host "Setting Execution Policy to Unrestricted..."
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted
}

#IMPORTING EOM MODULE, CONNECTING TO IT, CREATING A TEMP DIRECTORY AND EXPORTING SHARED MAILBOXES
Import-Module ExchangeOnlineManagement
Connect-ExchangeOnline
$test3142dir = Test-Path C:\by3142
$3142csv = "C:\by3142\SharedMailboxUsers"
$counter = [int]$filenumber = 1
$csv = ".csv"
$3142csvfullpath = "$3142csv"+"$counter"+"$csv"
if (-not $test3142dir) {
    New-Item -Path "c:\" -Name "by3142" -ItemType "directory"
    Get-Recipient -RecipientTypeDetails SharedMailbox -Resultsize unlimited | select PrimarySmtpAddress | export-csv $3142csvfullpath
} else {
    $test3142csv = Test-Path $3142csvfullpath
    if ($test3142csv -eq $false) {
        Get-Recipient -RecipientTypeDetails SharedMailbox -Resultsize unlimited | select PrimarySmtpAddress | export-csv $3142csvfullpath
    } else {
        while ($test3142csv) {
            $counter = $counter + 1
            $3142csvfullpath = "$3142csv"+"$counter"+"$csv"
            $test3142csv = Test-Path $3142csvfullpath
            if ($test3142csv -eq $false) {
                Get-Recipient -RecipientTypeDetails SharedMailbox -Resultsize unlimited | select PrimarySmtpAddress | export-csv $3142csvfullpath
            }
        }
    }
}

#MAKING SURE THE CSV FILE IS CORRECT, SETTING VARIABLES AND CONFIGURING THE MAILBOXES
Read-Host -Prompt "CHECK C:\BY3142 FOR THE SHARED MAILBOXES CSV FILE. PRESS ANY KEY TO CONTINUE IF THE CONTAININGS OF THE FILE ARE OKAY 1/2"
Read-Host -Prompt "CHECK C:\BY3142 FOR THE SHARED MAILBOXES CSV FILE. PRESS ANY KEY TO CONTINUE IF THE CONTAININGS OF THE FILE ARE OKAY 2/2"

$mailboxes = Import-Csv -Path $3142csvfullpath
$okmessage = "OK! `n"
$mailboxes | ForEach-Object {
    Write-Output "==================== `n"
    Write-Output $_.PrimarySmtpAddress
    set-mailbox $_.PrimarySmtpAddress -MessageCopyForSentAsEnabled $True
    set-mailbox $_.PrimarySmtpAddress -MessageCopyForSendOnBehalfEnabled $True
    Write-Host $okmessage -ForegroundColor Green
}
Stop-Transcript

Read-Host -Prompt "
___.           ________  ____   _____ ________  
\_ |__ ___.__. \_____  \/_   | /  |  |\_____  \ 
 | __ <   |  |   _(__  < |   |/   |  |_/  ____/ 
 | \_\ \___  |  /       \|   /    ^   /       \ 
 |___  / ____| /______  /|___\____   |\_______ \
     \/\/             \/          |__|        \/
       END OF SCRIPT. PRESS ENTER TO EXIT.       
     THE TRANSCRIPT CAN BE FOUND ON DESKTOP.
                        "
