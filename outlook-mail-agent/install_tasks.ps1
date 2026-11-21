<#
 Crée les 2 tâches planifiées Windows de l'agent (à lancer une fois, PowerShell, dossier du projet) :
   - MailAgent-Run     : relève + Excel + alertes, toutes les 30 min
   - MailAgent-Weekly  : reporting hebdomadaire, le lundi à 08:00
 Les tâches s'exécutent sous VOTRE compte, session ouverte (obligatoire pour Outlook COM).
 Usage :  powershell -ExecutionPolicy Bypass -File .\install_tasks.ps1 [-Python "C:\...\python.exe"]
#>
param([string]$Python = (Get-Command python).Source, [int]$EveryMinutes = 30)

$dir  = $PSScriptRoot
$user = "$env:USERDOMAIN\$env:USERNAME"
$prin = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive
$set  = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
          -ExecutionTimeLimit (New-TimeSpan -Hours 1)

$run = New-ScheduledTaskAction -Execute $Python -Argument "-m mailagent run -c config.yaml" -WorkingDirectory $dir
$trg = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes)
Register-ScheduledTask -TaskName "MailAgent-Run" -Action $run -Trigger $trg -Principal $prin -Settings $set -Force | Out-Null

$wk  = New-ScheduledTaskAction -Execute $Python -Argument "-m mailagent weekly-report -c config.yaml" -WorkingDirectory $dir
$trg = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 08:00
Register-ScheduledTask -TaskName "MailAgent-Weekly" -Action $wk -Trigger $trg -Principal $prin -Settings $set -Force | Out-Null

Write-Host "OK : tâches MailAgent-Run (toutes les $EveryMinutes min) et MailAgent-Weekly (lundi 08:00) créées."
Write-Host "Test immédiat : Start-ScheduledTask -TaskName MailAgent-Run   |   Logs : output\agent.log"
