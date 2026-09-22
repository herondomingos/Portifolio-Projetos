[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$SlackWebhookUrl = $env:SLACK_WEBHOOK_URL
if (-not $SlackWebhookUrl) { Write-Error "Variavel SLACK_WEBHOOK_URL nao definida"; exit 1 }

# Teste 1: Alerta de sucesso
$Body = @{ text = ":white_check_mark: *DNS Sync CNAME - OK*`nTotal: 45 | OK: 42 | Criados: 2 | Atualizados: 1`n• [CRIAR] api.example.com (CNAME) -> api.example.com.cdn.cloudflare.net`n• [ATUALIZAR] app.example.com (CNAME) -> app.example.com.cdn.cloudflare.net" } | ConvertTo-Json -Compress
Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $Body -ContentType "application/json"
Write-Output "Enviado: alerta de sucesso"

Start-Sleep -Seconds 2

# Teste 2: Alerta de erro
$Body = @{ text = ":rotating_light: *DNS Sync CNAME - ERRO*`n```Validacao pos-sync: 2 registro(s) com falha:`napi.example.com`napp.example.com```" } | ConvertTo-Json -Compress
Invoke-RestMethod -Uri $SlackWebhookUrl -Method Post -Body $Body -ContentType "application/json"
Write-Output "Enviado: alerta de erro"
