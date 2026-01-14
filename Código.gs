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
      sheet.appendRow(["Data de Entrega", "Data Recebimento", "Arquivo", "Cliente", "Marca", "Local Entrega", "Qtd", "Unidade", "Valor (R$)", "Ordem de Compra"]);
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
          p.valor,
          p.ordemCompra || "N/D"                     // Ordem de Compra
        ]);
      }
    }

    if (novasLinhas.length > 0) {
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 10).setValues(novasLinhas);
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

    var dados = sheet.getRange(inicio, 1, numLinhas, 10).getValues();
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
        formatarValor(row[8]),           // Valor (R$)
        row[9] ? row[9].toString() : ""  // Ordem de Compra
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

// ========================================
// SISTEMA DE FATURAMENTO
// ========================================

/**
 * Lê dados da aba "Dados1" (ordem de compra, valor, cliente)
 * @returns {Array} Array de objetos com os dados
 */
function lerDados1() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados1");
    if (!sheet) {
      Logger.log("⚠️ Aba 'Dados1' não encontrada");
      return [];
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("⚠️ Aba 'Dados1' vazia (sem dados além do cabeçalho)");
      return [];
    }

    // Pega dados a partir da linha 2 (pula cabeçalho)
    var dados = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    var resultado = [];
    dados.forEach(function(row) {
      if (row[0] && row[1]) { // Precisa ter pelo menos OC e Valor
        resultado.push({
          ordemCompra: row[0].toString().trim(),
          valor: typeof row[1] === 'number' ? row[1] : parseFloat(row[1]) || 0,
          cliente: row[2] ? row[2].toString().trim() : "Sem Cliente"
        });
      }
    });

    Logger.log("✅ Lidos " + resultado.length + " registros da aba Dados1");
    return resultado;
  } catch (erro) {
    Logger.log("❌ Erro ao ler Dados1: " + erro.toString());
    return [];
  }
}

/**
 * Busca a marca de uma OC na aba "Dados" (aba principal)
 * @param {string} oc - Ordem de Compra
 * @returns {string} Nome da marca ou "Sem Marca"
 */
function buscarMarcaPorOC(oc) {
  try {
    if (!oc) return "Sem Marca";

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
    if (!sheet) return "Sem Marca";

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return "Sem Marca";

    // Busca nas últimas 1000 linhas (otimização)
    var numLinhas = Math.min(1000, lastRow - 1);
    var inicio = lastRow - numLinhas + 1;

    // Pega apenas colunas OC (índice 10) e Marca (índice 5)
    var dados = sheet.getRange(inicio, 1, numLinhas, 10).getValues();

    // Percorre de trás pra frente (dados mais recentes primeiro)
    for (var i = dados.length - 1; i >= 0; i--) {
      var ocLinha = dados[i][9] ? dados[i][9].toString().trim() : ""; // Coluna J (índice 9)
      if (ocLinha === oc.toString().trim()) {
        var marca = dados[i][4] ? dados[i][4].toString().trim() : "Sem Marca"; // Coluna E (índice 4)
        return marca;
      }
    }

    return "Sem Marca";
  } catch (erro) {
    Logger.log("❌ Erro ao buscar marca da OC " + oc + ": " + erro.toString());
    return "Sem Marca";
  }
}

/**
 * Retorna pedidos a faturar (card 1)
 * Agrupa por cliente+marca, soma valores
 */
function getPedidosAFaturar() {
  try {
    Logger.log("📊 Iniciando getPedidosAFaturar...");

    var dados = lerDados1();

    if (dados.length === 0) {
      return {
        sucesso: true,
        timestamp: obterTimestamp(),
        dados: []
      };
    }

    // Agrupa por cliente+marca
    var agrupamentoMap = {};

    dados.forEach(function(item) {
      // Busca a marca pela OC
      var marca = buscarMarcaPorOC(item.ordemCompra);
      var chave = item.cliente + "|" + marca;

      if (!agrupamentoMap[chave]) {
        agrupamentoMap[chave] = {
          cliente: item.cliente,
          marca: marca,
          valor: 0
        };
      }

      agrupamentoMap[chave].valor += item.valor;
    });

    // Converte para array
    var resultado = Object.keys(agrupamentoMap).map(function(chave) {
      return agrupamentoMap[chave];
    });

    // Ordena por cliente (alfabético) e depois por valor (maior primeiro)
    resultado.sort(function(a, b) {
      if (a.cliente !== b.cliente) {
        return a.cliente.localeCompare(b.cliente);
      }
      return b.valor - a.valor;
    });

    Logger.log("✅ getPedidosAFaturar concluído: " + resultado.length + " linhas (cliente+marca)");

    return {
      sucesso: true,
      timestamp: obterTimestamp(),
      dados: resultado
    };

  } catch (erro) {
    Logger.log("❌ Erro em getPedidosAFaturar: " + erro.toString());
    return {
      sucesso: false,
      timestamp: obterTimestamp(),
      dados: [],
      erro: erro.toString()
    };
  }
}

/**
 * Sistema de snapshot para detectar faturamento
 * Salva snapshot atual e retorna o que foi faturado desde o último snapshot
 */
function getFaturamentoDia() {
  try {
    Logger.log("💰 Iniciando getFaturamentoDia...");

    var props = PropertiesService.getScriptProperties();
    var snapshotAnterior = props.getProperty('SNAPSHOT_DADOS1');
    var timestampAnterior = props.getProperty('SNAPSHOT_TIMESTAMP');

    // Lê estado atual
    var dadosAtuais = lerDados1();

    // Cria map do estado atual (OC -> dados)
    var mapaAtual = {};
    dadosAtuais.forEach(function(item) {
      mapaAtual[item.ordemCompra] = item;
    });

    var faturado = [];

    // Se não há snapshot anterior, cria o primeiro
    if (!snapshotAnterior) {
      Logger.log("📸 Criando primeiro snapshot...");
      props.setProperty('SNAPSHOT_DADOS1', JSON.stringify(mapaAtual));
      props.setProperty('SNAPSHOT_TIMESTAMP', obterTimestamp());

      return {
        sucesso: true,
        timestamp: timestampAnterior,
        dados: [],
        mensagem: "Primeiro snapshot criado. Aguardando próxima verificação."
      };
    }

    // Compara com snapshot anterior
    var mapaAnterior = JSON.parse(snapshotAnterior);

    Object.keys(mapaAnterior).forEach(function(oc) {
      var itemAnterior = mapaAnterior[oc];
      var itemAtual = mapaAtual[oc];

      var valorFaturado = 0;

      if (!itemAtual) {
        // OC sumiu = faturou tudo
        valorFaturado = itemAnterior.valor;
      } else if (itemAtual.valor < itemAnterior.valor) {
        // Valor diminuiu = faturou a diferença
        valorFaturado = itemAnterior.valor - itemAtual.valor;
      }

      if (valorFaturado > 0) {
        var marca = buscarMarcaPorOC(oc);

        faturado.push({
          cliente: itemAnterior.cliente,
          valor: valorFaturado,
          marca: marca,
          oc: oc
        });
      }
    });

    // Agrupa faturamento por cliente+marca
    var faturadoAgrupado = {};

    faturado.forEach(function(item) {
      var chave = item.cliente + "|" + item.marca;

      if (!faturadoAgrupado[chave]) {
        faturadoAgrupado[chave] = {
          cliente: item.cliente,
          marca: item.marca,
          valor: 0
        };
      }

      faturadoAgrupado[chave].valor += item.valor;
    });

    var resultado = Object.keys(faturadoAgrupado).map(function(chave) {
      return faturadoAgrupado[chave];
    });

    // Ordena por valor (maior primeiro)
    resultado.sort(function(a, b) {
      return b.valor - a.valor;
    });

    // Atualiza snapshot
    props.setProperty('SNAPSHOT_DADOS1', JSON.stringify(mapaAtual));
    props.setProperty('SNAPSHOT_TIMESTAMP', obterTimestamp());

    Logger.log("✅ getFaturamentoDia concluído: " + resultado.length + " itens faturados");

    return {
      sucesso: true,
      timestamp: timestampAnterior,
      dados: resultado
    };

  } catch (erro) {
    Logger.log("❌ Erro em getFaturamentoDia: " + erro.toString());
    return {
      sucesso: false,
      timestamp: null,
      dados: [],
      erro: erro.toString()
    };
  }
}

/**
 * Função auxiliar para obter timestamp formatado
 */
function obterTimestamp() {
  var agora = new Date();
  var dia = ("0" + agora.getDate()).slice(-2);
  var mes = ("0" + (agora.getMonth() + 1)).slice(-2);
  var ano = agora.getFullYear();
  var hora = ("0" + agora.getHours()).slice(-2);
  var min = ("0" + agora.getMinutes()).slice(-2);

  return dia + "/" + mes + "/" + ano + " às " + hora + ":" + min;
}

/**
 * Função manual para executar a verificação de faturamento
 * USE ESTA FUNÇÃO PARA EXECUTAR MANUALMENTE
 */
function executarVerificacaoFaturamento() {
  Logger.log("🔄 Executando verificação manual de faturamento...");

  var resultado = getFaturamentoDia();

  if (resultado.sucesso) {
    Logger.log("✅ Verificação concluída com sucesso!");
    Logger.log("📊 Itens faturados: " + resultado.dados.length);

    if (resultado.dados.length > 0) {
      Logger.log("💰 Detalhes do faturamento:");
      resultado.dados.forEach(function(item) {
        Logger.log("   - " + item.cliente + " (" + item.marca + "): R$ " + item.valor.toFixed(2));
      });
    } else {
      Logger.log("ℹ️ Nenhum faturamento detectado nesta verificação");
    }
  } else {
    Logger.log("❌ Erro na verificação: " + resultado.erro);
  }

  return resultado;
}

/**
 * Configura triggers automáticos (8h e 19h)
 * EXECUTE ESTA FUNÇÃO UMA VEZ PARA CONFIGURAR OS HORÁRIOS AUTOMÁTICOS
 */
function setupTriggers() {
  // Remove triggers antigos para evitar duplicação
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'executarVerificacaoFaturamento') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Cria trigger para 8h
  ScriptApp.newTrigger('executarVerificacaoFaturamento')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Cria trigger para 19h
  ScriptApp.newTrigger('executarVerificacaoFaturamento')
    .timeBased()
    .atHour(19)
    .everyDays(1)
    .create();

  Logger.log("✅ Triggers configurados com sucesso!");
  Logger.log("⏰ Verificações automáticas às 8h e 19h todos os dias");
}
