param(
    [Parameter(Mandatory=$true)]
    [string]$BackupFile,
    [switch]$DryRun
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$InternalZone = "example.com"

if (-not (Test-Path $BackupFile)) { Write-Error "Arquivo nao encontrado: $BackupFile"; exit 1 }

$Records = Import-Csv $BackupFile
Write-Output "Restaurando $(($Records).Count) registros de: $BackupFile`n"

if ($DryRun) { Write-Output "*** MODO DRY RUN - nenhuma alteracao sera feita ***`n" }

foreach ($R in $Records) {
    $Existing = Get-DnsServerResourceRecord -ZoneName $InternalZone -Name $R.HostName -ErrorAction SilentlyContinue |
        Where-Object { $_.RecordType -eq $R.RecordType }

    if ($R.RecordType -eq "A") {
        $CurrentValue = if ($Existing) { $Existing.RecordData.IPv4Address.ToString() } else { $null }

        if ($CurrentValue -eq $R.Data) {
            Write-Output "[OK] $($R.HostName) (A) -> $($R.Data)"
            continue
        }

        if ($DryRun) {
            $Action = if ($Existing) { "RESTAURAR" } else { "CRIAR" }
            Write-Output "[DRY RUN] $Action : $($R.HostName) (A) | ATUAL: $CurrentValue -> BACKUP: $($R.Data)"
            continue
        }

        if ($Existing) {
            $Old = $Existing.Clone()
            $Existing.RecordData.IPv4Address = [System.Net.IPAddress]::Parse($R.Data)
            Set-DnsServerResourceRecord -ZoneName $InternalZone -OldInputObject $Old -NewInputObject $Existing
        } else {
            Add-DnsServerResourceRecordA -ZoneName $InternalZone -Name $R.HostName -IPv4Address $R.Data
        }
        Write-Output "[RESTAURADO] $($R.HostName) (A) -> $($R.Data)"
    }
    elseif ($R.RecordType -eq "CNAME") {
        $CurrentValue = if ($Existing) { $Existing.RecordData.HostNameAlias.TrimEnd('.') } else { $null }

        if ($CurrentValue -eq $R.Data.TrimEnd('.')) {
            Write-Output "[OK] $($R.HostName) (CNAME) -> $($R.Data)"
            continue
        }

        if ($DryRun) {
            $Action = if ($Existing) { "RESTAURAR" } else { "CRIAR" }
            Write-Output "[DRY RUN] $Action : $($R.HostName) (CNAME) | ATUAL: $CurrentValue -> BACKUP: $($R.Data)"
            continue
        }

        if ($Existing) {
            Remove-DnsServerResourceRecord -ZoneName $InternalZone -Name $R.HostName -RRType CNAME -Force
        }
        Add-DnsServerResourceRecordCName -ZoneName $InternalZone -Name $R.HostName -HostNameAlias $R.Data
        Write-Output "[RESTAURADO] $($R.HostName) (CNAME) -> $($R.Data)"
    }
}

Write-Output "`nRestore finalizado."
