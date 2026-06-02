# Registra tarefa no Task Scheduler do Windows
# Roda sync_prazos.py a cada 30 minutos, para sempre, em segundo plano
# Execute este script UMA VEZ como administrador

$Python  = (Get-Command python).Source
$Script  = "C:\Users\ach\temp_mobile_fix\dashboard-juridico\sync_prazos.py"
$NomeTarefa = "SyncPrazosSupabase"

$action  = New-ScheduledTaskAction -Execute $Python -Argument $Script
$trigger = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 30) -Once -At (Get-Date)
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -StartWhenAvailable

Register-ScheduledTask -TaskName $NomeTarefa -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force

Write-Host "Tarefa '$NomeTarefa' registrada — sync a cada 30 minutos."
