// --- ARQUIVO: Código.gs ---

// 1. O SITE (Para o ser humano ver)
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
      .setTitle('Pedidos por Marca - Marfim Bahia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. A API (Para o Robô Python enviar dados)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("Dados"); // Certifique-se que o nome da aba é 'Dados'
    
    if (!sheet) {
      // Se não existir, cria e põe cabeçalho
      sheet = doc.insertSheet("Dados");
      sheet.appendRow(["Data de Entrega", "Data Recebimento", "Arquivo", "Cliente", "Marca", "Local Entrega", "Qtd", "Unidade", "Valor (R$)"]);
    }

    var json = JSON.parse(e.postData.contents);
    var lista = json.pedidos; // O Python manda { "pedidos": [...] }
    var novasLinhas = [];

    // Verificação simples de duplicidade (olhando ultimos 500 registros para ser rápido)
    var ultimaLinha = sheet.getLastRow();
    var arquivosExistentes = [];
    if (ultimaLinha > 1) {
      // Pega apenas a coluna C (Arquivo) - mudou de B para C por causa da nova coluna
      var dadosC = sheet.getRange(Math.max(2, ultimaLinha - 500), 3, Math.min(500, ultimaLinha-1), 1).getValues();
      arquivosExistentes = dadosC.map(function(r){ return r[0]; });
    }

    for (var i = 0; i < lista.length; i++) {
      var p = lista[i];
      if (arquivosExistentes.indexOf(p.arquivo) === -1) {
        novasLinhas.push([
          p.dataEntrega || p.dataPedido || p.data,  // Data de Entrega (aceita vários formatos)
          p.dataRecebimento || "",                   // Data Recebimento
          p.arquivo,
          p.cliente,
          p.marca,
          p.local,
          p.qtd,
          p.unidade,
          p.valor
        ]);
      }
    }

    if (novasLinhas.length > 0) {
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 9).setValues(novasLinhas);
      return ContentService.createTextOutput(JSON.stringify({"status":"Sucesso", "msg": novasLinhas.length + " novos."})).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status":"Neutro", "msg": "Sem novidades."})).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (erro) {
    return ContentService.createTextOutput(JSON.stringify({"status":"Erro", "msg": erro.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 3. FUNÇÃO QUE O SITE CHAMA PARA PEGAR DADOS DA PLANILHA
function getDadosPlanilha() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
    if (!sheet) {
      Logger.log("⚠️ Aba 'Dados' não encontrada");
      return [];
    }

    var lastRow = sheet.getLastRow();
    Logger.log("📊 Última linha: " + lastRow);

    if (lastRow < 2) {
      Logger.log("⚠️ Planilha vazia (sem dados além do cabeçalho)");
      return [];
    }

    // Pega até 1000 registros mais recentes para otimizar
    var numLinhas = Math.min(1000, lastRow - 1);
    var inicio = lastRow - numLinhas + 1;

    var dados = sheet.getRange(inicio, 1, numLinhas, 9).getValues();
    Logger.log("✅ Recuperados " + dados.length + " registros");

    // Formata os dados para garantir compatibilidade
    var dadosFormatados = dados.map(function(row) {
      return [
        formatarData(row[0]),            // Data de Entrega
        formatarData(row[1]),            // Data Recebimento
        row[2] ? row[2].toString() : "", // Arquivo
        row[3] ? row[3].toString() : "", // Cliente
        row[4] ? row[4].toString() : "", // Marca
        row[5] ? row[5].toString() : "", // Local Entrega
        formatarNumero(row[6]),          // Qtd
        row[7] ? row[7].toString() : "", // Unidade
        formatarValor(row[8])            // Valor (R$)
      ];
    });

    Logger.log("✅ Dados formatados com sucesso");
    return dadosFormatados;

  } catch (erro) {
    Logger.log("❌ Erro em getDadosPlanilha: " + erro.toString());
    throw new Error("Erro ao buscar dados: " + erro.message);
  }
}

// Funções auxiliares de formatação
function formatarData(valor) {
  if (!valor) return "";
  if (valor instanceof Date) {
    var dia = ("0" + valor.getDate()).slice(-2);
    var mes = ("0" + (valor.getMonth() + 1)).slice(-2);
    var ano = valor.getFullYear();
    return dia + "/" + mes + "/" + ano;
  }
  return valor.toString();
}

function formatarNumero(valor) {
  if (!valor) return "0";
  if (typeof valor === 'number') {
    return valor.toString();
  }
  return valor.toString();
}

function formatarValor(valor) {
  if (!valor) return "R$ 0,00";
  if (typeof valor === 'number') {
    return "R$ " + valor.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  }
  // Se já vier formatado, retorna como está
  return valor.toString();
}
