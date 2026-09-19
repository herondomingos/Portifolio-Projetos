# Gestão Automatizada de Domínios e Vencimentos

Automação em Google Apps Script que resolve um problema clássico de infraestrutura: domínios corporativos comprados em provedores diferentes, com vencimentos espalhados, e o risco real de perder um domínio crítico por falta de acompanhamento.

## O problema

Antes desta automação, o controle de vencimento de domínios era manual: planilhas desatualizadas, informação espalhada entre e-mails de diferentes registradores e nenhum alerta proativo antes do vencimento.

## A solução

O script roda periodicamente (via trigger do Apps Script) e:

1. **Consulta a API da GoDaddy** para trazer a data de expiração de todos os domínios ativos.
2. **Varre a caixa do Gmail** em busca de e-mails de confirmação da America Registry e do Registro.br, extraindo domínio e data via regex.
3. **Consolida tudo** em uma única planilha mestra, sem sobrescrever dados mais recentes por dados mais antigos (comparação de datas).
4. **Detecta e reporta mudanças** — domínios novos, renovados ou removidos — e envia um resumo automático para um canal do Slack via webhook.

Em produção, gerencia atualmente mais de 500 domínios corporativos e cresce automaticamente a cada novo domínio adquirido, sem intervenção manual.

## Stack

- Google Apps Script (JavaScript)
- Google Sheets API
- Gmail API (parsing de e-mails de fornecedores)
- API da GoDaddy
- Slack Incoming Webhooks

## Configuração

Este repositório **não contém nenhuma credencial**. Antes de rodar, configure em *Project Settings > Script Properties* do seu projeto Apps Script:

| Propriedade | Descrição |
|---|---|
| `GODADDY_KEY` | Chave da API GoDaddy |
| `GODADDY_SECRET` | Segredo da API GoDaddy |
| `SLACK_WEBHOOK` | URL do Incoming Webhook do Slack |
| `PLANILHA_ID` | ID da planilha do Google Sheets usada como banco de dados |

> Nunca commite essas credenciais no código-fonte — use sempre `PropertiesService`.
