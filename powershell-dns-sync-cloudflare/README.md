# Sincronismo Automatizado de DNS (Cloudflare ➔ Windows Server)

Automação em PowerShell projetada para manter a infraestrutura de DNS interno (Windows Server) sempre sincronizada com os apontamentos públicos gerenciados na nuvem (Cloudflare), resolvendo a falta de paridade entre os ambientes.

## O problema

Manter a paridade de registros DNS entre a nuvem pública e o ambiente interno on-premise era um desafio manual. Alterações ou novos apontamentos criados no Cloudflare precisavam ser replicados manualmente no Windows Server, o que gerava lentidão na entrega de serviços, risco de falha humana (typos) e total falta de histórico ou alerta rápido sobre o que havia sido modificado na infraestrutura.

## A solução

O conjunto de scripts roda periodicamente e de forma totalmente autônoma realizando as seguintes etapas:

1. **Backup preventivo:** Antes de qualquer modificação, faz um dump dos registros DNS atuais do Windows Server para um arquivo CSV (permitindo rollback imediato com o script de restore).
2. **Consulta inteligente:** Consome a API v4 do Cloudflare com paginação para trazer o estado oficial da zona pública.
3. **Sincronização e Correção:** Compara a nuvem com o servidor local. Cria ou atualiza os registros divergentes. O script possui inteligência para tratar limitações do protocolo DNS, como converter automaticamente registros CNAME no Apex do domínio para registros do tipo A.
4. **Validação:** Resolve o nome recém-criado no servidor local (127.0.0.1) para garantir que a propagação interna ocorreu com sucesso.
5. **Auditoria e Alertas:** Gera uma página HTML local estática que serve como um painel de monitoramento e envia o resumo das operações (OK, erros, criados, atualizados) para um canal de infraestrutura no Slack.

## Stack

* PowerShell (Automação e interação com o módulo DnsServer)
* Windows Server (Microsoft DNS)
* API do Cloudflare v4 (Integração de zonas e registros)
* Slack Incoming Webhooks (Alertas de monitoramento)
* HTML/CSS (Geração do Dashboard de Status)

## Configuração

Este repositório **não contém nenhuma credencial**. Antes de rodar, configure as seguintes variáveis de ambiente no Windows Server (ou no seu cofre de senhas):
* `SLACK_WEBHOOK_URL`: URL do webhook para envio dos resumos.
* `CF_API_TOKEN`: Token gerado no Cloudflare com permissão de leitura de zona DNS.