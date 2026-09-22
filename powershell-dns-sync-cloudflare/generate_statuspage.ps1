$LogDir = "$PSScriptRoot\Logs"
$OutputFile = "$PSScriptRoot\status.html"

$Logs = Get-ChildItem "$LogDir\dns-sync-*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 20

$Rows = foreach ($Log in $Logs) {
    $Content = Get-Content $Log.FullName -Raw
    $Status = if ($Content -match "ERRO|FALHA|Exception") { "<span style='color:red'>&#x2718; ERRO</span>" } else { "<span style='color:green'>&#x2714; OK</span>" }
    $Date = $Log.LastWriteTime.ToString("dd/MM/yyyy HH:mm:ss")
    $Type = if ($Log.Name -match "CNAME") { "CNAME" } else { "A/AAAA" }

    # Extrair resumo
    $Resumo = ""
    if ($Content -match "Total: .+") { $Resumo = $Matches[0] }

    "<tr><td>$Date</td><td>$Type</td><td>$Status</td><td>$Resumo</td><td><a href='Logs/$($Log.Name)'>ver log</a></td></tr>"
}

$Html = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="60">
    <title>DNS Sync Status - $($env:COMPUTERNAME)</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background: #000000; color: #fff; }
        h1 { color: #0078D7; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #333; padding: 10px; text-align: left; }
        th { background: #1a1a1a; color: #0078D7; }
        tr:nth-child(even) { background: #1a1a1a; }
        tr:hover { background: #2a2a2a; }
        a { color: #0078D7; }
        .header { display: flex; justify-content: space-between; align-items: center; }
        .updated { color: #888; font-size: 12px; }
        .logo { font-weight: bold; font-size: 24px; color: #0078D7; }
    </style>
</head>
<body>
    <div class="header">
        <h1>&#x1F310; DNS Sync Status | <span class="logo">Empresa Corp</span></h1>
        <span class="updated">Servidor: $($env:COMPUTERNAME) | Atualizado: $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | Auto-refresh: 60s</span>
    </div>
    <table>
        <tr><th>Data/Hora</th><th>Tipo</th><th>Status</th><th>Resumo</th><th>Log</th></tr>
        $($Rows -join "`n        ")
    </table>
</body>
</html>
"@

$Html | Out-File -FilePath $OutputFile -Encoding UTF8
Write-Output "Pagina gerada: $OutputFile"
