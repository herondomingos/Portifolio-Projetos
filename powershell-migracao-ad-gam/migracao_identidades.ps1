<#
.SYNOPSIS
    Migração em massa de identidades corporativas (Active Directory + Google Workspace).
.DESCRIPTION
    Usado durante um rebranding corporativo para atualizar UPN e endereço de
    e-mail de um lote de usuários simultaneamente no Active Directory e no
    Google Workspace (via GAM), com log de auditoria completo de sucessos e
    falhas em cada sistema.

    VERSÃO SANITIZADA PARA PORTFÓLIO — a lista de usuários é apenas
    ilustrativa. Em produção, a lista real (150 contas, no caso do projeto
    original) era carregada de uma fonte externa (CSV ou planilha), nunca
    hardcoded no script.
.NOTES
    Pré-requisitos: módulo ActiveDirectory do RSAT e GAM (Google Apps
    Manager) instalado e autenticado no host que executa o script.
#>

param(
    # Permite carregar a lista de um CSV externo em vez de usar o
    # array de exemplo abaixo. CSV esperado com colunas: OldEmail,NewEmail
    [string]$CaminhoCsv
)

Import-Module ActiveDirectory

$LogFile = "C:\Logs\Migracao_Identidades_Log.txt"
if (-not (Test-Path "C:\Logs")) { New-Item -ItemType Directory -Path "C:\Logs" | Out-Null }

# Lista de exemplo (dados fictícios) — substitua por -CaminhoCsv em produção
$UsersToMigrate = if ($CaminhoCsv) {
    Import-Csv -Path $CaminhoCsv
} else {
    @(
        [pscustomobject]@{ OldEmail = 'joao.silva@empresa-antiga.com.br'; NewEmail = 'joao.silva@novamarca.com' }
        [pscustomobject]@{ OldEmail = 'maria.souza@empresa-antiga.com.br'; NewEmail = 'maria.souza@novamarca.com' }
        [pscustomobject]@{ OldEmail = 'pedro.santos@empresa-antiga.com.br'; NewEmail = 'pedro.santos@novamarca.com' }
    )
}

Add-Content -Path $LogFile -Value "=== INÍCIO DA MIGRAÇÃO EM MASSA: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ==="
$Total = $UsersToMigrate.Count
$Contador = 1

foreach ($User in $UsersToMigrate) {
    $OldEmail = $User.OldEmail
    $NewEmail = $User.NewEmail

    Write-Host "[$Contador/$Total] Processando $OldEmail para $NewEmail..." -ForegroundColor Cyan

    # --- 1. ATUALIZAÇÃO NO ACTIVE DIRECTORY ---
    try {
        $AdUser = Get-ADUser -Filter "mail -eq '$OldEmail'"
        if ($AdUser) {
            Set-ADUser -Identity $AdUser.DistinguishedName -UserPrincipalName $NewEmail -Replace @{mail = $NewEmail}
            Add-Content -Path $LogFile -Value "[$OldEmail] [SUCESSO AD] Atualizado para $NewEmail."
        } else {
            Add-Content -Path $LogFile -Value "[$OldEmail] [ERRO AD] Usuário não encontrado no AD."
            Write-Host "  -> [ERRO AD] Usuário não encontrado no AD." -ForegroundColor Yellow
        }
    } catch {
        $CleanError = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
        Add-Content -Path $LogFile -Value "[$OldEmail] [ERRO AD] $CleanError"
        Write-Host "  -> [ERRO AD] Falha na atualização do AD." -ForegroundColor Red
    }

    # --- 2. ATUALIZAÇÃO NO GAM (GOOGLE WORKSPACE) ---
    $GamOutput = gam update user $OldEmail email $NewEmail 2>&1
    if ($LASTEXITCODE -eq 0) {
        Add-Content -Path $LogFile -Value "[$OldEmail] [SUCESSO GAM] Atualizado no Google Workspace."
    } else {
        $CleanErrorGAM = ($GamOutput -join " ") -replace "`r`n", " " -replace "`n", " "
        Add-Content -Path $LogFile -Value "[$OldEmail] [ERRO GAM] Detalhes: $CleanErrorGAM"
        Write-Host "  -> [ERRO GAM] Falha no GAM (Verifique o log)." -ForegroundColor Red
    }

    $Contador++
}

Add-Content -Path $LogFile -Value "=== FIM DA MIGRAÇÃO EM MASSA: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss') ==="
Write-Host "Migração concluída! Foram processados $Total usuários." -ForegroundColor Green
Write-Host "O relatório completo de execução está em: $LogFile" -ForegroundColor Cyan
