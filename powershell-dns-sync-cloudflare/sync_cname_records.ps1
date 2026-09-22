param(
    [switch]$DryRun
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- LOCK FILE: evita execucao paralela ---
$LockFile = "$PSScriptRoot\sync-cname.lock"
if (Test-Path $LockFile) {
    $LockAge = (Get-Date) - (Get-Item $LockFile).LastWriteTime
    if ($LockAge.TotalMinutes -lt 30) {
        Write-Error "Outra execucao em andamento (lock: $LockFile). Abortando."
        exit 1
    }
    Remove-Item $LockFile -Force
}
New-Item -Path $LockFile -ItemType File -Force | Out-Null

try {

# --- FUNCAO DE ALERTA SLACK ---
$SlackWebhookUrl = $env:SLACK_WEBHOOK_URL
$ServerName = $env:COMPUTERNAME
$Timestamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
function Send-SlackAlert($Message) {
    if (-not $SlackWebhookUrl) { return }
    $Body = @{ text = ":rotating_light: *DNS Sync CNAME - ERRO* [$ServerName] - $Timestamp`n```$Message```" } | ConvertTo-Json -Compress
    try { Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $Body -ContentType "application/json" } catch {}
}

$CloudflareApiToken = $env:CF_API_TOKEN
if (-not $CloudflareApiToken) { Send-SlackAlert "Variavel CF_API_TOKEN nao definida"; Write-Error "Variavel CF_API_TOKEN nao definida"; exit 1 }

$InternalZone = "example.com"
$MaxBackups = 30
$MaxLogs = 30
$Headers = @{ "Authorization" = "Bearer $CloudflareApiToken"; "Content-Type" = "application/json" }
$ZoneId = (Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones?name=$InternalZone" -Headers $Headers).result[0].id
if (-not $ZoneId) { Write-Error "Zone ID nao encontrado para $InternalZone"; exit 1 }
$BackupDir = "$PSScriptRoot\Backups"

New-Item -ItemType Directory -Path "$PSScriptRoot\Logs" -Force | Out-Null
Start-Transcript -Path "$PSScriptRoot\Logs\dns-sync-CNAME-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

if ($DryRun) { Write-Output "*** MODO DRY RUN - nenhuma alteracao sera feita ***`n" }

# Backup registros CNAME e A do apex atuais
New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
$BackupFile = "$BackupDir\dns-backup-CNAME-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
Get-DnsServerResourceRecord -ZoneName $InternalZone | Where-Object { $_.RecordType -in @("CNAME","A") } |
    Select-Object HostName, RecordType, @{N='Data';E={if($_.RecordData.IPv4Address){$_.RecordData.IPv4Address}else{$_.RecordData.HostNameAlias}}}, TimeToLive |
    Export-Csv -Path $BackupFile -NoTypeInformation -Encoding UTF8
Write-Output "Backup salvo: $BackupFile`n"
Get-ChildItem "$BackupDir\dns-backup-CNAME-*.csv" | Sort-Object CreationTime -Descending | Select-Object -Skip $MaxBackups | Remove-Item -Force

# --- PAGINACAO: busca todos os registros CNAME ---
$AllRecords = @()
$Page = 1
do {
    $Response = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones/$ZoneId/dns_records?type=CNAME&per_page=100&page=$Page" -Headers $Headers
    if (-not $Response.success) {
        Write-Error "Falha ao consultar Cloudflare: $($Response.errors | ConvertTo-Json)"
        Stop-Transcript; exit 1
    }
    $AllRecords += $Response.result
    $TotalPages = $Response.result_info.total_pages
    $Page++
} while ($Page -le $TotalPages)

Write-Output "Total de registros CNAME encontrados: $($AllRecords.Count)`n"

# --- SYNC ---
$SyncResults = @()

foreach ($Record in $AllRecords) {
    $HostName = $Record.name -replace "\.$([regex]::Escape($InternalZone))$", ""
    if ($HostName -eq $Record.name) { $HostName = "@" }
    $IsApex = ($HostName -eq "@")

    # Ignorar subdomínios de múltiplos níveis
    if (-not $IsApex -and $HostName -match "\.") {
        Write-Output "[IGNORADO] $($Record.name) - subdominio multinivel, requer zona delegada"
        continue
    }

    # --- APEX: cria como registro A (CNAME nao permitido no apex) ---
    if ($IsApex) {
        $ResolveTarget = if ($Record.proxied) { $Record.name } else { $Record.content }
        $ResolvedIPs = [System.Net.Dns]::GetHostAddresses($ResolveTarget) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }

        foreach ($IP in $ResolvedIPs) {
            $Existing = Get-DnsServerResourceRecord -ZoneName $InternalZone -Name "@" -RRType A -ErrorAction SilentlyContinue |
                Where-Object { $_.RecordData.IPv4Address.ToString() -eq $IP.ToString() }

            if ($Existing) {
                Write-Output "[OK] $($Record.name) (A - apex) -> $($IP.ToString())"
                $SyncResults += @{ Name=$Record.name; Type="A"; Value=$IP.ToString(); Status="OK" }
                continue
            }

            if ($DryRun) {
                Write-Output "[DRY RUN] CRIAR : $($Record.name) (A - apex) -> $($IP.ToString())"
                continue
            }

            Add-DnsServerResourceRecordA -ZoneName $InternalZone -Name "@" -IPv4Address $IP.ToString() -TimeToLive ([TimeSpan]::FromSeconds(300))
            Write-Output "[CRIADO] $($Record.name) (A - apex) -> $($IP.ToString())"
            $SyncResults += @{ Name=$Record.name; Type="A"; Value=$IP.ToString(); Status="CRIADO" }
        }
        continue
    }

    # --- SUBDOMINIO: CNAME com .cdn.cloudflare.net se proxy ativo ---
    $Target = if ($Record.proxied) { "$($Record.name).cdn.cloudflare.net" } else { $Record.content }

    $Existing = Get-DnsServerResourceRecord -ZoneName $InternalZone -Name $HostName -ErrorAction SilentlyContinue |
        Where-Object { $_.RecordType -eq "CNAME" }

    if ($Existing -and $Existing.RecordData.HostNameAlias.TrimEnd('.') -eq $Target.TrimEnd('.')) {
        Write-Output "[OK] $($Record.name) (CNAME) -> $Target"
        $SyncResults += @{ Name=$Record.name; Type="CNAME"; Value=$Target; Status="OK" }
        continue
    }

    $Action = if ($Existing) { "ATUALIZAR" } else { "CRIAR" }
    $CurrentValue = if ($Existing) { $Existing.RecordData.HostNameAlias.TrimEnd('.') } else { "N/A" }

    if ($DryRun) {
        if ($Action -eq "ATUALIZAR") {
            Write-Output "[DRY RUN] $Action : $($Record.name) (CNAME) | ATUAL: $CurrentValue -> NOVO: $Target"
        } else {
            Write-Output "[DRY RUN] $Action : $($Record.name) (CNAME) -> $Target"
        }
        continue
    }

    if ($Existing) {
        Remove-DnsServerResourceRecord -ZoneName $InternalZone -Name $HostName -RRType CNAME -Force
    }

    try {
        Add-DnsServerResourceRecordCName -ZoneName $InternalZone -Name $HostName -HostNameAlias $Target
        if ($Action -eq "ATUALIZAR") {
            Write-Output "[ATUALIZADO] $($Record.name) (CNAME) | ANTERIOR: $CurrentValue -> NOVO: $Target"
        } else {
            Write-Output "[CRIADO] $($Record.name) (CNAME) -> $Target"
        }
        $SyncResults += @{ Name=$Record.name; Type="CNAME"; Value=$Target; Status=$Action }
    } catch {
        Write-Output "[ERRO] $($Record.name) (CNAME) -> $Target | $($_.Exception.Message)"
        $SyncResults += @{ Name=$Record.name; Type="CNAME"; Value=$Target; Status="ERRO" }
    }
}

# --- RESUMO E ALERTA ---
$CountOK = ($SyncResults | Where-Object { $_.Status -eq "OK" }).Count
$CountCriado = ($SyncResults | Where-Object { $_.Status -eq "CRIAR" }).Count
$CountAtualizado = ($SyncResults | Where-Object { $_.Status -eq "ATUALIZAR" }).Count
$CountErro = ($SyncResults | Where-Object { $_.Status -eq "ERRO" }).Count

Write-Output "`n--- RESUMO ---"
Write-Output "Total: $($SyncResults.Count) | OK: $CountOK | Criados: $CountCriado | Atualizados: $CountAtualizado | Erros: $CountErro"

if ($CountErro -gt 0) {
    Write-Output "`n--- REGISTROS COM ERRO ---"
    $SyncResults | Where-Object { $_.Status -eq "ERRO" } | ForEach-Object {
        Write-Output "  - $($_.Name) ($($_.Type)) -> $($_.Value)"
    }
}

# --- VALIDACAO POS-SYNC ---
$Errors = 0
$FailedNames = @()
if (-not $DryRun -and ($CountCriado + $CountAtualizado) -gt 0) {
    Write-Output "`n--- VALIDACAO POS-SYNC ---"
    foreach ($R in $SyncResults) {
        if ($R.Status -eq "OK") { continue }
        $Resolved = Resolve-DnsName -Name $R.Name -Server "127.0.0.1" -ErrorAction SilentlyContinue
        if ($Resolved) {
            Write-Output "[VALIDADO] $($R.Name) -> responde no DNS interno"
        } else {
            Write-Output "[FALHA] $($R.Name) -> NAO responde no DNS interno"
            $FailedNames += $R.Name
            $Errors++
        }
    }
}

# Alerta Slack com resumo
if (-not $DryRun) {
    if ($Errors -gt 0) {
        $FailedList = $FailedNames -join "`n"
        Send-SlackAlert "Validacao pos-sync: $Errors registro(s) com falha:`n$FailedList"
    } else {
        $ChangedRecords = $SyncResults | Where-Object { $_.Status -ne "OK" }
        $Details = ""
        if ($ChangedRecords.Count -gt 0) {
            $Details = "`n" + (($ChangedRecords | ForEach-Object { "• [$($_.Status)] $($_.Name) ($($_.Type)) -> $($_.Value)" }) -join "`n")
        }
        $Summary = ":white_check_mark: *DNS Sync CNAME - OK* [$ServerName] - $Timestamp`nTotal: $($SyncResults.Count) | OK: $CountOK | Criados: $CountCriado | Atualizados: $CountAtualizado$Details"
        if ($SlackWebhookUrl) {
            $Body = @{ text = $Summary } | ConvertTo-Json -Compress
            try { Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $Body -ContentType "application/json" } catch {}
        }
    }
}

Stop-Transcript

# Rotacao de logs
Get-ChildItem "$PSScriptRoot\Logs\dns-sync-CNAME-*.log" | Sort-Object CreationTime -Descending | Select-Object -Skip $MaxLogs | Remove-Item -Force

} catch {
    Send-SlackAlert $_.Exception.Message
    Write-Error $_.Exception.Message
} finally {
    # Remove lock file
    Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
}
