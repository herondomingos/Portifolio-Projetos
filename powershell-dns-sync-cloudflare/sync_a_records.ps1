param(
    [switch]$DryRun
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- LOCK FILE: evita execucao paralela ---
$LockFile = "$PSScriptRoot\sync-a.lock"
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
    $Body = @{ text = ":rotating_light: *DNS Sync A/AAAA - ERRO* [$ServerName] - $Timestamp`n```$Message```" } | ConvertTo-Json -Compress
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
Start-Transcript -Path "$PSScriptRoot\Logs\dns-sync-A-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

if ($DryRun) { Write-Output "*** MODO DRY RUN - nenhuma alteracao sera feita ***`n" }

# Backup registros A e AAAA atuais
New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
$BackupFile = "$BackupDir\dns-backup-A-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"
Get-DnsServerResourceRecord -ZoneName $InternalZone | Where-Object { $_.RecordType -in @("A","AAAA") } |
    Select-Object HostName, RecordType, @{N='Data';E={if($_.RecordData.IPv4Address){$_.RecordData.IPv4Address}else{$_.RecordData.IPv6Address}}}, TimeToLive |
    Export-Csv -Path $BackupFile -NoTypeInformation -Encoding UTF8
Write-Output "Backup salvo: $BackupFile`n"
Get-ChildItem "$BackupDir\dns-backup-A-*.csv" | Sort-Object CreationTime -Descending | Select-Object -Skip $MaxBackups | Remove-Item -Force

# --- PAGINACAO: busca todos os registros A e AAAA ---
$AllRecords = @()
foreach ($Type in @("A","AAAA")) {
    $Page = 1
    do {
        $Resp = Invoke-RestMethod -Uri "https://api.cloudflare.com/client/v4/zones/$ZoneId/dns_records?type=$Type&per_page=100&page=$Page" -Headers $Headers
        if ($Resp.success) { $AllRecords += $Resp.result }
        $TotalPages = $Resp.result_info.total_pages
        $Page++
    } while ($Page -le $TotalPages)
}

Write-Output "Total de registros A/AAAA encontrados: $($AllRecords.Count)`n"

# --- SYNC ---
$SyncResults = @()

foreach ($Record in $AllRecords) {
    $HostName = $Record.name -replace "\.$([regex]::Escape($InternalZone))$", ""
    if ($HostName -eq $Record.name) { $HostName = "@" }

    # Ignorar apex - tratado pelo script CNAME
    if ($HostName -eq "@") { continue }

    # Ignorar subdomínios de múltiplos níveis
    if ($HostName -match "\.") {
        Write-Output "[IGNORADO] $($Record.name) - subdominio multinivel, requer zona delegada"
        continue
    }

    # Se proxy ativo, resolve o IP publico (Cloudflare)
    if ($Record.proxied) {
        $ResolvedIPs = [System.Net.Dns]::GetHostAddresses($Record.name) | Where-Object { $_.AddressFamily -eq 'InterNetwork' }
        if ($ResolvedIPs) { $Record.content = $ResolvedIPs[0].ToString() }
    }

    $Existing = Get-DnsServerResourceRecord -ZoneName $InternalZone -Name $HostName -ErrorAction SilentlyContinue |
        Where-Object { $_.RecordType -eq $Record.type } | Select-Object -First 1

    # Verifica se ja esta correto
    $CurrentValue = $null
    if ($Existing) {
        if ($Record.type -eq "A" -and $Existing.RecordData.PSObject.Properties['IPv4Address']) {
            $CurrentValue = $Existing.RecordData.IPv4Address.ToString()
        } elseif ($Record.type -eq "AAAA" -and $Existing.RecordData.PSObject.Properties['IPv6Address']) {
            $CurrentValue = $Existing.RecordData.IPv6Address.ToString()
        }
    }
    if ($Existing -and $CurrentValue -eq $Record.content) {
        Write-Output "[OK] $($Record.name) ($($Record.type)) -> $($Record.content)"
        $SyncResults += @{ Name=$Record.name; Type=$Record.type; Value=$Record.content; Status="OK" }
        continue
    }

    $Action = if ($Existing) { "ATUALIZAR" } else { "CRIAR" }
    $ProxyTag = if ($Record.proxied) { " [PROXY]" } else { "" }

    if ($DryRun) {
        if ($Action -eq "ATUALIZAR") {
            Write-Output "[DRY RUN] $Action : $($Record.name) ($($Record.type)) | ATUAL: $CurrentValue -> NOVO: $($Record.content)$ProxyTag"
        } else {
            Write-Output "[DRY RUN] $Action : $($Record.name) ($($Record.type)) -> $($Record.content)$ProxyTag"
        }
        continue
    }

    try {
        if ($Record.type -eq "A") {
            if ($Existing) {
                $Old = $Existing.Clone()
                $Existing.RecordData.IPv4Address = [System.Net.IPAddress]::Parse($Record.content)
                Set-DnsServerResourceRecord -ZoneName $InternalZone -OldInputObject $Old -NewInputObject $Existing
            } else {
                Add-DnsServerResourceRecordA -ZoneName $InternalZone -Name $HostName -IPv4Address $Record.content -TimeToLive ([TimeSpan]::FromSeconds($Record.ttl))
            }
        }
        elseif ($Record.type -eq "AAAA") {
            if ($Existing) {
                $Old = $Existing.Clone()
                $Existing.RecordData.IPv6Address = [System.Net.IPAddress]::Parse($Record.content)
                Set-DnsServerResourceRecord -ZoneName $InternalZone -OldInputObject $Old -NewInputObject $Existing
            } else {
                Add-DnsServerResourceRecordAAAA -ZoneName $InternalZone -Name $HostName -IPv6Address $Record.content -TimeToLive ([TimeSpan]::FromSeconds($Record.ttl))
            }
        }

        if ($Action -eq "ATUALIZAR") {
            Write-Output "[ATUALIZADO] $($Record.name) ($($Record.type)) | ANTERIOR: $CurrentValue -> NOVO: $($Record.content)$ProxyTag"
        } else {
            Write-Output "[CRIADO] $($Record.name) ($($Record.type)) -> $($Record.content)$ProxyTag"
        }
        $SyncResults += @{ Name=$Record.name; Type=$Record.type; Value=$Record.content; Status=$Action }
    } catch {
        Write-Output "[ERRO] $($Record.name) ($($Record.type)) -> $($Record.content) | $($_.Exception.Message)"
        $SyncResults += @{ Name=$Record.name; Type=$Record.type; Value=$Record.content; Status="ERRO" }
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
        $Summary = ":white_check_mark: *DNS Sync A/AAAA - OK* [$ServerName] - $Timestamp`nTotal: $($SyncResults.Count) | OK: $CountOK | Criados: $CountCriado | Atualizados: $CountAtualizado$Details"
        if ($SlackWebhookUrl) {
            $Body = @{ text = $Summary } | ConvertTo-Json -Compress
            try { Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $Body -ContentType "application/json" } catch {}
        }
    }
}

Stop-Transcript

# Rotacao de logs
Get-ChildItem "$PSScriptRoot\Logs\dns-sync-A-*.log" | Sort-Object CreationTime -Descending | Select-Object -Skip $MaxLogs | Remove-Item -Force

} catch {
    Send-SlackAlert $_.Exception.Message
    Write-Error $_.Exception.Message
} finally {
    # Remove lock file
    Remove-Item $LockFile -Force -ErrorAction SilentlyContinue
}
