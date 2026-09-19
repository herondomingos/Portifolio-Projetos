# Migração em Massa de Identidades (Active Directory + Google Workspace)

Script em PowerShell usado durante um rebranding corporativo para atualizar UPN e endereço de e-mail de contas de usuário simultaneamente em dois sistemas — Active Directory e Google Workspace — de forma auditável e com tratamento de erro independente por sistema.

## O problema

Trocar o domínio de e-mail da empresa (rebranding) exige atualizar o UPN e o atributo `mail` no Active Directory *e* o e-mail correspondente no Google Workspace, para cada conta. Feito manualmente, é lento e sujeito a erro humano — e uma falha parcial (AD atualizado, Google não, ou vice-versa) é difícil de rastrear depois.

## A solução

O script processa uma lista de usuários (`OldEmail` → `NewEmail`) e, para cada um:

1. Localiza o usuário no AD pelo atributo `mail` atual.
2. Atualiza o UPN e o atributo `mail` via `Set-ADUser`.
3. Atualiza o e-mail correspondente no Google Workspace via GAM (Google Apps Manager).
4. Registra o resultado de **cada sistema separadamente** em um log com timestamp — se o AD funcionar mas o GAM falhar (ou vice-versa), fica claro no log qual conta precisa de reprocessamento manual.

Usado em produção para migrar 150 contas de usuário durante um rebranding corporativo, sem impacto em ferramentas de terceiros.

## Stack

- PowerShell + módulo ActiveDirectory (RSAT)
- GAM (Google Apps Manager) para Google Workspace

## Uso

```powershell
# Com lista de exemplo embutida (para teste)
.\migracao_identidades.ps1

# Com lista real via CSV (colunas: OldEmail,NewEmail)
.\migracao_identidades.ps1 -CaminhoCsv .\usuarios_para_migrar.csv
```

> Este repositório não contém nenhum dado real de usuário — a lista de exemplo no script usa e-mails fictícios.
