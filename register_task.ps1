$botDir = "C:\Users\User\OneDrive - Частная компания Adal Commodities Ltd\Рабочий стол\Мне\My_sturtup\TelegramBot\opt\asudkr-bot"
$cmdFile = Join-Path $botDir "run_bot.cmd"

Unregister-ScheduledTask -TaskName "ASUDKR-TelegramBot" -Confirm:$false -ErrorAction SilentlyContinue

$action   = New-ScheduledTaskAction -Execute "cmd.exe" -Argument ('/c "' + $cmdFile + '"') -WorkingDirectory $botDir
$trigger  = New-ScheduledTaskTrigger -AtLogOn -User "$env:USERDOMAIN\$env:USERNAME"
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -Hidden -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName "ASUDKR-TelegramBot" -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "ASU DKR Telegram bot (nakladnye -> Excel)" | Out-Null

Write-Host "Задача создана. Запускаю сейчас..."
Start-ScheduledTask -TaskName "ASUDKR-TelegramBot"
Start-Sleep -Seconds 3
Get-ScheduledTask -TaskName "ASUDKR-TelegramBot" | Get-ScheduledTaskInfo | Format-List TaskName, LastRunTime, LastTaskResult, NextRunTime

Write-Host "Готово. Можно закрыть это окно."
