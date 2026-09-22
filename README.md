# Projetos de Automação — Heron Domingos

Automações desenvolvidas para resolver problemas reais de infraestrutura de TI corporativa (Smart Fit), com foco em eliminar trabalho manual repetitivo e reduzir risco operacional.

Todo o código aqui foi sanitizado para portfólio público: credenciais, IDs de planilha, e-mails e nomes de domínio reais foram removidos ou substituídos por exemplos fictícios. A lógica e a arquitetura são as mesmas usadas em produção.

## Projetos

### [`apps-script-dominios/`](./apps-script-dominios)
Sistema que consolida dados de 3 fontes (API GoDaddy, e-mails da America Registry e do Registro.br) para monitorar o vencimento de domínios corporativos, com alertas automáticos via Slack. Em produção, gerencia mais de 500 domínios e cresce automaticamente a cada nova compra.

### [`powershell-migracao-ad-gam/`](./powershell-migracao-ad-gam)
Script de migração em massa de identidades (Active Directory + Google Workspace) usado durante um rebranding corporativo — 150 contas de usuário migradas com log de auditoria por sistema.

### [`python-retrofit-flow/`](./python-retrofit-flow)
Aplicação desktop que automatiza o preenchimento de planilhas de controle de ativos a partir de checklists de retrofit em campo. Reduziu o tempo de processamento de um ciclo de ~800 horas para menos de 15 minutos.


[`powershell-dns-sync-cloudflare/`](powershell-dns-sync-cloudflare/)

Conjunto de scripts que automatiza a sincronização de registros DNS (A, AAAA, CNAME) entre a nuvem pública (Cloudflare) e a infraestrutura interna (Windows Server). Conta com sistema integrado de backup/rollback em CSV, validação pós-sincronismo, dashboard HTML auto-atualizável e envio de alertas de auditoria via Slack.

## Contato

- LinkedIn: [linkedin.com/in/heron-domingos](https://www.linkedin.com/in/heron-domingos-4b1a0b164/)
- E-mail: heron_domingos@outlook.com
