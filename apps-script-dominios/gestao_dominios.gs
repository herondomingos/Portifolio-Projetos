/**
 * Gestão Automatizada de Domínios e Vencimentos
 * -----------------------------------------------
 * Consolida dados de 3 fontes (API GoDaddy, e-mails de confirmação da
 * America Registry e do Registro.br) em uma planilha mestra e envia
 * alertas de renovação/vencimento para o Slack via webhook.
 *
 * VERSÃO SANITIZADA PARA PORTFÓLIO — todas as credenciais, IDs de
 * planilha e nomes de domínio reais foram removidos. Antes de usar,
 * configure as variáveis abaixo via
 * Project Settings > Script Properties (nunca deixe hardcoded).
 */

// --- CONFIGURAÇÃO (lida de Script Properties, nunca hardcoded) ---
const PROPS = PropertiesService.getScriptProperties();
const GODADDY_KEY = PROPS.getProperty('GODADDY_KEY');
const GODADDY_SECRET = PROPS.getProperty('GODADDY_SECRET');
const SLACK_WEBHOOK = PROPS.getProperty('SLACK_WEBHOOK');
const PLANILHA_ID = PROPS.getProperty('PLANILHA_ID');

function atualizarPlanilhaMestra() {
  var spreadsheet = SpreadsheetApp.openById(PLANILHA_ID);
  var sheet = spreadsheet.getSheets()[0];

  var dadosDominios = {};
  var datasAntigas = {};

  // 1. PRESERVAR DADOS EXISTENTES E MAPEAMENTO ANTIGO
  var ultimaLinhaOriginal = sheet.getLastRow();
  if (ultimaLinhaOriginal > 1) {
    var dadosAtuais = sheet.getRange(2, 1, ultimaLinhaOriginal - 1, 3).getValues();
    dadosAtuais.forEach(function(r) {
      var d = String(r[0]).toLowerCase().trim();
      if (ehDominioValido(d)) {
        var dataPlanilha = r[1];
        if (dataPlanilha instanceof Date) {
          dataPlanilha = Utilities.formatDate(dataPlanilha, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        dadosDominios[d] = { data: dataPlanilha, provedor: r[2] };
        datasAntigas[d] = dataPlanilha;
      }
    });
  }

  // 1.5. LISTA DE CONTROLE EM CASO DE FALHA DE API
  // Exemplo de estrutura — em produção, mantenha essa lista fora do
  // código-fonte (ex: em outra aba da própria planilha ou em
  // Script Properties como JSON) para não expor domínios da empresa
  // em um repositório público.
  var dominiosFixos = {
    "exemplo-dominio.com.br": "17/06/2027",
    "outro-exemplo.com": "12/08/2027"
  };

  for (var df in dominiosFixos) {
    var prov = df.endsWith(".br") ? "Registro.br" : "America Registry";

    if (!dadosDominios[df] || dadosDominios[df].data === "Verificar") {
      dadosDominios[df] = { data: dominiosFixos[df], provedor: prov };
    } else {
      var pFixa = dominiosFixos[df].split('/');
      var pAtual = dadosDominios[df].data.split('/');
      if (pFixa.length === 3 && pAtual.length === 3) {
        var numFixa = parseInt(pFixa[2] + pFixa[1] + pFixa[0], 10);
        var numAtual = parseInt(pAtual[2] + pAtual[1] + pAtual[0], 10);
        if (numFixa > numAtual) {
          dadosDominios[df] = { data: dominiosFixos[df], provedor: prov };
        }
      }
    }
  }

  // 2. BUSCA NA API DA GODADDY
  try {
    var gdDomains = getGoDaddyDomainsAPI();
    gdDomains.forEach(function(item) {
      dadosDominios[item.domain] = { data: item.expires, provedor: "GoDaddy" };
    });
  } catch (e) { Logger.log("Erro GoDaddy: " + e); }

  // 3. BUSCA AMERICA REGISTRY (Gmail)
  var queryAR = 'in:anywhere from:americaregistry.com ("Auto Renew Notification" OR "Order Confirmation" OR "Domain Renewal Confirmation" OR "Domain Registration Certificate") newer_than:120d';
  GmailApp.search(queryAR, 0, 150).forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody().replace(/\s+/g, " ");

      var domM = body.match(/domain name\s+([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i) ||
                 body.match(/DOMAIN CREATE:\s*([a-zA-Z0-9.-]+)/i) ||
                 body.match(/DOMAIN RENEW:\s*([a-zA-Z0-9.-]+)/i);

      var dateM = body.match(/shortly on\s+(\d{4}-\d{2}-\d{2})/i);
      var orderDateM = body.match(/Order Confirmation on\s+(\d{4}-\d{2}-\d{2})/i);
      var termM = body.match(/Term:\s*(\d+)\s*YEAR/i);

      var renewConfirmDomM = body.match(/advise that\s+([a-zA-Z0-9.-]+)\s+has been/i) ||
                             body.match(/expiry renewal date for\s+([a-zA-Z0-9.-]+)\s+is/i);
      var renewConfirmDateM = body.match(/expiry renewal date for.*?is\s+(\d{4}-\d{2}-\d{2})/i);

      if ((!domM || !ehDominioValido(domM[1])) && renewConfirmDomM) {
        domM = renewConfirmDomM;
      }
      if (!dateM && renewConfirmDateM) {
        dateM = renewConfirmDateM;
      }

      var dataFinal = "Verificar";

      if (dateM) {
        dataFinal = formatarDataBR(dateM[1]);
      } else if (orderDateM && termM) {
        var dataCriacao = new Date(orderDateM[1] + "T12:00:00Z");
        var anosTermo = parseInt(termM[1], 10);
        dataCriacao.setFullYear(dataCriacao.getFullYear() + anosTermo);
        dataFinal = Utilities.formatDate(dataCriacao, Session.getScriptTimeZone(), "dd/MM/yyyy");
      } else if (orderDateM) {
        dataFinal = formatarDataBR(orderDateM[1]) + " (Criado)";
      }

      if (domM && ehDominioValido(domM[1])) {
        var d = domM[1].toLowerCase().trim();

        var salvar = true;
        if (dadosDominios[d] && dadosDominios[d].data && dadosDominios[d].data !== "Verificar" && dataFinal !== "Verificar") {
          var pNovo = dataFinal.split('/');
          var pAtual = dadosDominios[d].data.split('/');

          if (pNovo.length === 3 && pAtual.length === 3) {
            var numNovo = parseInt(pNovo[2] + pNovo[1] + pNovo[0], 10);
            var numAtual = parseInt(pAtual[2] + pAtual[1] + pAtual[0], 10);

            if (numNovo <= numAtual) {
              salvar = false;
            }
          }
        }

        if (salvar) {
          dadosDominios[d] = {
            data: dataFinal !== "Verificar" ? dataFinal : (dadosDominios[d] ? dadosDominios[d].data : "Verificar"),
            provedor: "America Registry"
          };
        }
      }
    });
  });

  // 4. BUSCA REGISTRO.BR (Gmail)
  // Ajuste o remetente/domínio de e-mail conforme o seu ambiente.
  var queryBR = 'in:anywhere "registro.br" "Domínio:" ("Manutenção" OR "Expiração" OR "Período Contratado") newer_than:60d';
  var threadsBR = GmailApp.search(queryBR, 0, 100);

  threadsBR.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      var body = msg.getPlainBody().replace(/\s+/g, " ");
      var domM = body.match(/Dom[íi]nio:\s*([a-zA-Z0-9.-]+\.br)/i);

      var periodoM = body.match(/Per[íi]odo Contratado:.*?at[ée]\s+(\d{2}\/\d{2}\/\d{4})/i);
      var manutencaoM = body.match(/Manuten[çc][ãa]o\s+de\s+(\d{2}\/\d{2}\/\d{4})\s+a\s+(\d{2}\/\d{2}\/\d{4})/i);
      var expM = body.match(/Expira[çc][ãa]o:\s*(\d{2}\/\d{2}\/\d{4})/i);

      var dateVal = null;
      if (periodoM) {
        dateVal = periodoM[1];
      } else if (manutencaoM) {
        dateVal = manutencaoM[2];
      } else if (expM) {
        dateVal = expM[1];
      }

      if (domM && dateVal && ehDominioValido(domM[1])) {
        var dBR = domM[1].toLowerCase().trim();

        var salvar = true;
        if (dadosDominios[dBR] && dadosDominios[dBR].data && dadosDominios[dBR].data !== "Verificar") {
          var pNovo = dateVal.split('/');
          var pAtual = dadosDominios[dBR].data.split('/');

          if (pNovo.length === 3 && pAtual.length === 3) {
            var numNovo = parseInt(pNovo[2] + pNovo[1] + pNovo[0], 10);
            var numAtual = parseInt(pAtual[2] + pAtual[1] + pAtual[0], 10);

            if (numNovo <= numAtual) {
              salvar = false;
            }
          }
        }

        if (salvar) {
          dadosDominios[dBR] = { data: dateVal, provedor: "Registro.br" };
        }
      }
    });
  });

  // 5. COMPARAÇÃO (NOVIDADES, REMOÇÕES E RENOVAÇÕES)
  var listaFinal = [];
  var adicionados = [];
  var removidos = [];
  var renovados = [];

  for (var d in dadosDominios) {
    if (ehDominioValido(d)) {
      listaFinal.push([d, dadosDominios[d].data, dadosDominios[d].provedor]);

      if (!datasAntigas[d]) {
        adicionados.push(d);
      } else if (datasAntigas[d] !== dadosDominios[d].data) {
        renovados.push(d + " (de `" + datasAntigas[d] + "` para `" + dadosDominios[d].data + "`)");
      }
    }
  }

  for (var dAntigo in datasAntigas) {
    if (!dadosDominios[dAntigo]) {
      removidos.push(dAntigo);
    }
  }

  // 6. REESCRITA E ORGANIZAÇÃO ALFABÉTICA
  sheet.clear();
  var headers = [["Domínio", "Data de Vencimento", "Provedor"]];
  sheet.getRange(1, 1, 1, 3).setValues(headers).setFontWeight("bold").setBackground("#f3f3f3");

  if (listaFinal.length > 0) {
    listaFinal.sort(function(a, b) {
      return a[0].toLowerCase() < b[0].toLowerCase() ? -1 : 1;
    });
    sheet.getRange(2, 1, listaFinal.length, 3).setValues(listaFinal);
  }
  sheet.autoResizeColumns(1, 3);

  // 7. ENVIAR NOTIFICAÇÃO PARA O SLACK
  var urlPlanilha = spreadsheet.getUrl();
  enviarNotificacaoSlack(listaFinal.length, adicionados, removidos, renovados, urlPlanilha);
}

// --- FUNÇÃO DO SLACK ---
function enviarNotificacaoSlack(totalDominios, adicionados, removidos, renovados, urlPlanilha) {
  if (!SLACK_WEBHOOK || SLACK_WEBHOOK === '') return;
  var dataAtual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy 'às' HH:mm");

  var textoSlack = "✅ *Automação de Domínios Concluída*\n\nO script de controle de domínios foi executado.\n\n📅 *Data:* `" + dataAtual + "`\n🌐 *Total de domínios:* `" + totalDominios + "`\n📊 *Planilha:* <" + urlPlanilha + "|Clique aqui para acessar>\n";

  if (adicionados.length > 0) {
    textoSlack += "\n🟢 *Novos domínios adicionados (" + adicionados.length + "):*\n• " + adicionados.join("\n• ");
  }

  if (renovados.length > 0) {
    textoSlack += "\n🔵 *Domínios renovados/atualizados (" + renovados.length + "):*\n• " + renovados.join("\n• ");
  }

  if (removidos.length > 0) {
    textoSlack += "\n🔴 *Domínios removidos (" + removidos.length + "):*\n• " + removidos.join("\n• ");
  }

  if (adicionados.length === 0 && removidos.length === 0 && renovados.length === 0) {
    textoSlack += "\n⚪ *Nenhuma alteração:* A lista de domínios e vencimentos continua idêntica à última verificação.";
  }

  var payload = { "text": textoSlack };
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    UrlFetchApp.fetch(SLACK_WEBHOOK, options);
  } catch (e) {
    Logger.log("Erro ao enviar para o Slack: " + e);
  }
}

// --- VALIDADOR ANTI-POLUIÇÃO ---
function ehDominioValido(texto) {
  if (!texto) return false;
  var d = texto.toLowerCase().trim();
  if (d.includes(" ") || !d.includes(".") || d.startsWith(".") || d.includes("$") || d.includes(":") || d.includes(",")) return false;

  // Lista de padrões que costumam poluir os resultados de regex
  // em e-mails de fornecedores (rastreadores, links de imagem, etc).
  var blacklist = [
    "autorenew.link", "link.click", "logogd.link", "none.link", "ocp.email"
  ];
  if (blacklist.indexOf(d) !== -1 || d.startsWith("2f")) return false;
  if (d.split(".")[0].length < 2) return false;

  return true;
}

function getGoDaddyDomainsAPI() {
  var url = "https://api.godaddy.com/v1/domains?statuses=ACTIVE";
  var options = { "method": "get", "headers": { "Authorization": "sso-key " + GODADDY_KEY + ":" + GODADDY_SECRET }, "muteHttpExceptions": true };
  var response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) return [];
  var json = JSON.parse(response.getContentText());
  return json.map(function(item) {
    return { domain: item.domain.toLowerCase(), expires: Utilities.formatDate(new Date(item.expires), Session.getScriptTimeZone(), "dd/MM/yyyy") };
  });
}

function formatarDataBR(iso) {
  if (!iso || iso === "") return "";
  var p = iso.split('-');
  return p[2] + '/' + p[1] + '/' + p[0];
}
