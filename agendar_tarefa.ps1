# ================================================
#   AGENDADOR DE TAREFAS DO WINDOWS - v2
# ================================================
# Execute com: PowerShell -ExecutionPolicy Bypass -File agendar_tarefa.ps1

$NomeTarefa   = "AutomacaoPowerBI"
$PastaScript  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptPython = Join-Path $PastaScript "automacao_powerbi.py"
$PythonExe    = (Get-Command python -ErrorAction SilentlyContinue).Source

if (-not $PythonExe) {
    Write-Host "[ERRO] Python nao encontrado." -ForegroundColor Red
    exit 1
}

Write-Host "Criando tarefa agendada '$NomeTarefa'..." -ForegroundColor Cyan

Unregister-ScheduledTask -TaskName $NomeTarefa -Confirm:$false -ErrorAction SilentlyContinue

$Acao = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "`"$ScriptPython`"" `
    -WorkingDirectory $PastaScript

$Gatilho = New-ScheduledTaskTrigger -AtLogOn

$Config = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 0) `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $NomeTarefa `
    -Action $Acao `
    -Trigger $Gatilho `
    -Settings $Config `
    -RunLevel Highest `
    -Description "Automacao Power BI + Excel v2" | Out-Null

Write-Host "[OK] Tarefa criada! A automacao iniciara com o Windows." -ForegroundColor Green

$r = Read-Host "Iniciar agora? (S/N)"
if ($r -eq "S" -or $r -eq "s") {
    Start-ScheduledTask -TaskName $NomeTarefa
    Write-Host "[OK] Automacao iniciada em segundo plano!" -ForegroundColor Green
}
Read-Host "Pressione Enter para fechar"
