// --- ARQUIVO: Código.gs ---

// ========================================
// SISTEMA DE AUTENTICAÇÃO
// ========================================

/**
 * Verifica login contra a aba "senha" da planilha
 * @param {string} usuario - Nome de usuário
 * @param {string} senha - Senha
 * @returns {Object} Resultado da verificação
 */
function verificarLogin(usuario, senha) {
  try {
    Logger.log("🔐 Verificando login para usuário: " + usuario);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("senha");

    if (!sheet) {
      Logger.log("❌ Aba 'senha' não encontrada!");
      return {
        sucesso: false,
        mensagem: "Erro de configuração do sistema"
      };
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log("❌ Nenhum usuário cadastrado");
      return {
        sucesso: false,
        mensagem: "Nenhum usuário cadastrado"
      };
    }

    // Lê todos os usuários (pula cabeçalho)
    var dados = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    // Verifica se usuário e senha conferem
    for (var i = 0; i < dados.length; i++) {
      var usuarioNaAba = dados[i][0] ? dados[i][0].toString().trim().toUpperCase() : "";
      var senhaNaAba = dados[i][1] ? dados[i][1].toString().trim() : "";

      var usuarioDigitado = usuario ? usuario.toString().trim().toUpperCase() : "";
      var senhaDigitada = senha ? senha.toString().trim() : "";

      if (usuarioNaAba === usuarioDigitado && senhaNaAba === senhaDigitada) {
        Logger.log("✅ Login bem-sucedido para: " + usuario);

        // Gera token de sessão
        var token = gerarTokenSessao(usuario);

        return {
          sucesso: true,
          mensagem: "Login realizado com sucesso!",
          usuario: usuario,
          token: token
        };
      }
    }

    Logger.log("❌ Credenciais inválidas para: " + usuario);
    return {
      sucesso: false,
      mensagem: "Usuário ou senha incorretos"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao verificar login: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro ao verificar credenciais"
    };
  }
}

/**
 * Gera um token de sessão simples
 * @param {string} usuario - Nome de usuário
 * @returns {string} Token de sessão
 */
function gerarTokenSessao(usuario) {
  var agora = new Date().getTime();
  var props = PropertiesService.getScriptProperties();

  // Token = base64(usuario:timestamp)
  var tokenData = usuario + ":" + agora;
  var token = Utilities.base64Encode(tokenData);

  // Salva o token com timestamp
  props.setProperty('TOKEN_' + token, JSON.stringify({
    usuario: usuario,
    timestamp: agora
  }));

  Logger.log("🔑 Token gerado para: " + usuario);
  return token;
}

/**
 * Valida um token de sessão
 * @param {string} token - Token a validar
 * @returns {Object} Resultado da validação
 */
function validarToken(token) {
  try {
    if (!token) {
      return { valido: false, mensagem: "Token não fornecido" };
    }

    var props = PropertiesService.getScriptProperties();
    var tokenData = props.getProperty('TOKEN_' + token);

    if (!tokenData) {
      return { valido: false, mensagem: "Token inválido" };
    }

    var dados = JSON.parse(tokenData);
    var agora = new Date().getTime();
    var tempoDecorrido = agora - dados.timestamp;

    // Token válido por 8 horas (28800000 ms)
    var VALIDADE_TOKEN = 8 * 60 * 60 * 1000;

    if (tempoDecorrido > VALIDADE_TOKEN) {
      // Token expirado
      props.deleteProperty('TOKEN_' + token);
      return { valido: false, mensagem: "Sessão expirada" };
    }

    return {
      valido: true,
      usuario: dados.usuario
    };

  } catch (erro) {
    Logger.log("❌ Erro ao validar token: " + erro.toString());
    return { valido: false, mensagem: "Erro na validação" };
  }
}

/**
 * Faz logout invalidando o token
 * @param {string} token - Token a invalidar
 */
function fazerLogout(token) {
  try {
    if (token) {
      var props = PropertiesService.getScriptProperties();
      props.deleteProperty('TOKEN_' + token);
      Logger.log("👋 Logout realizado");
    }
    return { sucesso: true };
  } catch (erro) {
    Logger.log("❌ Erro ao fazer logout: " + erro.toString());
    return { sucesso: false };
  }
}

// 1. O SITE (Para o ser humano ver)
function doGet(e) {
  Logger.log("📄 doGet chamado - URL: " + (e ? JSON.stringify(e.parameter) : "sem parâmetros"));

  // Verifica se há token na URL
  var token = e && e.parameter ? e.parameter.token : null;

  if (token) {
    Logger.log("🔑 Token recebido na URL: " + token.substring(0, 10) + "...");

    // Valida o token
    var validacao = validarToken(token);
    Logger.log("✅ Validação do token: " + JSON.stringify(validacao));

    if (validacao.valido) {
      // Token válido - mostra a página principal
      Logger.log("✅ Token válido! Usuário: " + validacao.usuario);
      Logger.log("📄 Carregando Index.html para usuário: " + validacao.usuario);

      var template = HtmlService.createTemplateFromFile('Index');
      template.usuarioLogado = validacao.usuario;
      template.tokenSessao = token;

      Logger.log("📄 Template configurado com usuário: " + template.usuarioLogado);

      return template.evaluate()
          .setTitle('Pedidos por Marca - Marfim Bahia')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } else {
      Logger.log("❌ Token inválido ou expirado: " + validacao.mensagem);
    }
  } else {
    Logger.log("⚠️ Nenhum token fornecido na URL");
  }

  // Sem token ou token inválido - mostra página de login
  Logger.log("🔐 Mostrando página de login");
  var template = HtmlService.createTemplateFromFile('Login');
  return template.evaluate()
      .setTitle('Login - Marfim Bahia')
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
 * Cria um mapa de OC -> Marca carregando TODAS as linhas de uma vez (OTIMIZADO)
 * @returns {Object} Mapa com OC como chave e marca como valor
 */
function criarMapaOCMarca() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
    if (!sheet) {
      Logger.log("⚠️ Aba 'Dados' não encontrada");
      return {};
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("⚠️ Aba 'Dados' vazia");
      return {};
    }

    // Carrega TODAS as linhas (sem limite)
    var numLinhas = lastRow - 1;
    Logger.log("📥 Carregando mapa OC->Marca de TODAS as " + numLinhas + " linhas...");

    // Pega apenas as colunas necessárias: Marca (E/5) e OC (J/10)
    var dados = sheet.getRange(2, 1, numLinhas, 10).getValues();

    var mapa = {};
    var contador = 0;

    // Percorre e cria o mapa
    dados.forEach(function(row) {
      var oc = row[9] ? row[9].toString().trim() : ""; // Coluna J (índice 9)
      var marca = row[4] ? row[4].toString().trim() : "Sem Marca"; // Coluna E (índice 4)

      if (oc && oc !== "") {
        // Sobrescreve se já existe (pega a mais recente)
        mapa[oc] = marca;
        contador++;
      }
    });

    Logger.log("✅ Mapa criado com " + Object.keys(mapa).length + " OCs únicas de " + numLinhas + " linhas");
    return mapa;

  } catch (erro) {
    Logger.log("❌ Erro ao criar mapa OC->Marca: " + erro.toString());
    return {};
  }
}

/**
 * Busca a marca de uma OC no mapa pré-carregado
 * @param {string} oc - Ordem de Compra
 * @param {Object} mapaOCMarca - Mapa de OC->Marca
 * @returns {string} Nome da marca ou "Sem Marca"
 */
function buscarMarcaNoMapa(oc, mapaOCMarca) {
  if (!oc || !mapaOCMarca) return "Sem Marca";
  var ocLimpa = oc.toString().trim();
  return mapaOCMarca[ocLimpa] || "Sem Marca";
}

/**
 * Retorna pedidos a faturar (card 1) - OTIMIZADO
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

    Logger.log("📦 " + dados.length + " registros lidos da aba Dados1");

    // OTIMIZAÇÃO: Carrega todas as marcas de UMA VEZ
    var mapaOCMarca = criarMapaOCMarca();

    // Agrupa por cliente+marca
    var agrupamentoMap = {};

    dados.forEach(function(item) {
      // Busca a marca no mapa (rápido - O(1))
      var marca = buscarMarcaNoMapa(item.ordemCompra, mapaOCMarca);
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
 * Sistema de snapshot para detectar faturamento - OTIMIZADO
 * Salva snapshot atual e retorna o que foi faturado desde o último snapshot
 * IMPORTANTE: Só atualiza snapshot quando chamado via trigger (não na webapp)
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

    // OTIMIZAÇÃO: Carrega mapa de marcas UMA VEZ
    var mapaOCMarca = criarMapaOCMarca();

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
        // Busca marca no mapa (rápido)
        var marca = buscarMarcaNoMapa(oc, mapaOCMarca);

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

    // CORREÇÃO CRÍTICA: Atualiza snapshot SOMENTE via trigger, nunca via webapp
    // Isso evita que chamadas manuais destruam a detecção de faturamento
    // O snapshot só deve ser atualizado DEPOIS que o faturamento foi processado
    Logger.log("📸 Atualizando snapshot após detecção de faturamento...");
    props.setProperty('SNAPSHOT_DADOS1', JSON.stringify(mapaAtual));
    props.setProperty('SNAPSHOT_TIMESTAMP', obterTimestamp());

    // === LÓGICA ACUMULATIVA: Acumula faturamentos do mesmo dia ===
    var dataAtual = new Date();
    var diaAtual = ("0" + dataAtual.getDate()).slice(-2) + "/" +
                   ("0" + (dataAtual.getMonth() + 1)).slice(-2) + "/" +
                   dataAtual.getFullYear();

    var diaArmazenado = props.getProperty('FATURAMENTO_DATA');
    var faturamentoAcumulado = [];

    // Verifica se é um novo dia
    if (diaArmazenado !== diaAtual) {
      // Novo dia - reseta o acumulado
      Logger.log("📅 Novo dia detectado (" + diaAtual + ") - resetando acumulado de faturamento");
      props.setProperty('FATURAMENTO_DATA', diaAtual);
      faturamentoAcumulado = [];
    } else {
      // Mesmo dia - carrega o acumulado existente
      var ultimoFaturamento = props.getProperty('ULTIMO_FATURAMENTO');
      if (ultimoFaturamento) {
        faturamentoAcumulado = JSON.parse(ultimoFaturamento);
        Logger.log("📊 Mesmo dia - carregando acumulado existente (" + faturamentoAcumulado.length + " itens)");
      }
    }

    // Se houve novo faturamento nesta verificação, acumula com o existente
    if (resultado.length > 0) {
      Logger.log("💰 Novo faturamento detectado: " + resultado.length + " itens");

      // Cria mapa para acumular
      var mapAcumulado = {};

      // Primeiro, adiciona o que já estava acumulado
      faturamentoAcumulado.forEach(function(item) {
        var chave = item.cliente + "|" + item.marca;
        mapAcumulado[chave] = {
          cliente: item.cliente,
          marca: item.marca,
          valor: item.valor
        };
      });

      // Depois, soma o novo faturamento
      resultado.forEach(function(item) {
        var chave = item.cliente + "|" + item.marca;
        if (!mapAcumulado[chave]) {
          mapAcumulado[chave] = {
            cliente: item.cliente,
            marca: item.marca,
            valor: 0
          };
        }
        mapAcumulado[chave].valor += item.valor;
      });

      // Converte de volta para array
      var novoAcumulado = Object.keys(mapAcumulado).map(function(chave) {
        return mapAcumulado[chave];
      });

      // Ordena por valor (maior primeiro)
      novoAcumulado.sort(function(a, b) {
        return b.valor - a.valor;
      });

      // Salva o acumulado
      props.setProperty('ULTIMO_FATURAMENTO', JSON.stringify(novoAcumulado));
      props.setProperty('ULTIMO_FATURAMENTO_TIMESTAMP', obterTimestamp());

      Logger.log("💾 Salvou faturamento acumulado: " + novoAcumulado.length + " itens (cliente+marca)");

      // Atualiza resultado para retornar o acumulado
      resultado = novoAcumulado;

      // Salva no histórico da planilha (apenas quando é novo faturamento no acumulado)
      salvarFaturamentoNoHistorico(novoAcumulado, diaAtual);
    } else if (faturamentoAcumulado.length > 0) {
      // Não houve novo faturamento, mas há acumulado do dia
      Logger.log("ℹ️ Nenhum novo faturamento nesta verificação, mantendo acumulado do dia");
      resultado = faturamentoAcumulado;
    }

    Logger.log("✅ getFaturamentoDia concluído: " + resultado.length + " itens no total do dia");

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
 * Retorna o último faturamento detectado (para exibir na webapp)
 * Esta função NÃO recalcula, apenas retorna o que foi salvo
 */
function getUltimoFaturamento() {
  try {
    var props = PropertiesService.getScriptProperties();
    var ultimoFaturamento = props.getProperty('ULTIMO_FATURAMENTO');
    var timestamp = props.getProperty('ULTIMO_FATURAMENTO_TIMESTAMP');

    if (!ultimoFaturamento) {
      return {
        sucesso: true,
        timestamp: null,
        dados: [],
        mensagem: "Nenhum faturamento detectado ainda. Aguardando primeira verificação."
      };
    }

    var dados = JSON.parse(ultimoFaturamento);

    return {
      sucesso: true,
      timestamp: timestamp,
      dados: dados
    };

  } catch (erro) {
    Logger.log("❌ Erro em getUltimoFaturamento: " + erro.toString());
    return {
      sucesso: false,
      timestamp: null,
      dados: [],
      erro: erro.toString()
    };
  }
}

/**
 * Salva o faturamento do dia no histórico da planilha
 * @param {Array} dados - Array com os dados do faturamento
 * @param {string} data - Data no formato DD/MM/AAAA
 */
function salvarFaturamentoNoHistorico(dados, data) {
  try {
    if (!dados || dados.length === 0) {
      Logger.log("⚠️ Nenhum dado para salvar no histórico");
      return;
    }

    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("HistoricoFaturamento");

    // Cria a aba se não existir
    if (!sheet) {
      Logger.log("📋 Criando aba 'HistoricoFaturamento'...");
      sheet = doc.insertSheet("HistoricoFaturamento");
      // Adiciona cabeçalho
      sheet.appendRow(["Data", "Cliente", "Marca", "Valor Faturado", "Timestamp"]);
      // Formata cabeçalho
      var headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setBackground("#d32f2f");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var timestamp = obterTimestamp();
    var novasLinhas = [];

    // Verifica se já existe entrada para esta data
    var lastRow = sheet.getLastRow();
    var datasExistentes = [];

    if (lastRow > 1) {
      // Pega as datas já registradas (coluna A)
      var dadosExistentes = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      datasExistentes = dadosExistentes.map(function(row) { return row[0]; });
    }

    // Se já existe entrada para hoje, atualiza (substitui) ao invés de duplicar
    var jaExisteHoje = datasExistentes.indexOf(data) !== -1;

    if (jaExisteHoje) {
      Logger.log("🔄 Atualizando faturamento existente para " + data);

      // Remove linhas antigas do dia
      for (var i = lastRow; i >= 2; i--) {
        var dataLinha = sheet.getRange(i, 1).getValue();
        if (dataLinha === data) {
          sheet.deleteRow(i);
        }
      }
    }

    // Adiciona novas linhas
    dados.forEach(function(item) {
      novasLinhas.push([
        data,
        item.cliente,
        item.marca,
        item.valor,
        timestamp
      ]);
    });

    if (novasLinhas.length > 0) {
      var ultimaLinha = sheet.getLastRow();
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 5).setValues(novasLinhas);

      // Formata valores como moeda
      var valorRange = sheet.getRange(ultimaLinha + 1, 4, novasLinhas.length, 1);
      valorRange.setNumberFormat("R$ #,##0.00");

      Logger.log("✅ Salvou " + novasLinhas.length + " linhas no histórico para " + data);
    }

  } catch (erro) {
    Logger.log("❌ Erro ao salvar no histórico: " + erro.toString());
  }
}

/**
 * Retorna o histórico completo de faturamentos salvos na planilha
 * @returns {Object} Objeto com array de histórico
 */
function getHistoricoFaturamento() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");

    if (!sheet) {
      Logger.log("⚠️ Aba 'HistoricoFaturamento' não encontrada");
      return {
        sucesso: true,
        dados: [],
        mensagem: "Nenhum histórico disponível ainda."
      };
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log("⚠️ Histórico vazio");
      return {
        sucesso: true,
        dados: [],
        mensagem: "Nenhum histórico disponível ainda."
      };
    }

    // Lê todos os dados (pula cabeçalho)
    var dados = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

    var historico = [];

    dados.forEach(function(row) {
      historico.push({
        data: row[0].toString(),
        cliente: row[1].toString(),
        marca: row[2].toString(),
        valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0,
        timestamp: row[4].toString()
      });
    });

    // Ordena por data (mais recente primeiro)
    historico.sort(function(a, b) {
      // Converte DD/MM/AAAA para comparação
      var partesA = a.data.split('/');
      var partesB = b.data.split('/');
      var dataA = new Date(partesA[2], partesA[1] - 1, partesA[0]);
      var dataB = new Date(partesB[2], partesB[1] - 1, partesB[0]);
      return dataB - dataA;
    });

    Logger.log("✅ Retornou " + historico.length + " registros do histórico");

    return {
      sucesso: true,
      dados: historico
    };

  } catch (erro) {
    Logger.log("❌ Erro ao ler histórico: " + erro.toString());
    return {
      sucesso: false,
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
 * Função para resetar manualmente o acumulado de faturamento do dia
 * USE ESTA FUNÇÃO PARA LIMPAR/RESETAR O ACUMULADO (útil para testes ou ajustes)
 */
function resetarAcumuladoFaturamento() {
  Logger.log("🔄 Resetando acumulado de faturamento...");

  var props = PropertiesService.getScriptProperties();

  // Remove os dados acumulados
  props.deleteProperty('ULTIMO_FATURAMENTO');
  props.deleteProperty('ULTIMO_FATURAMENTO_TIMESTAMP');
  props.deleteProperty('FATURAMENTO_DATA');

  Logger.log("✅ Acumulado resetado com sucesso!");
  Logger.log("ℹ️ Na próxima verificação, o acumulado começará do zero");

  return {
    sucesso: true,
    mensagem: "Acumulado resetado com sucesso"
  };
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

// ========================================
// FUNÇÕES DE TESTE E DEBUG
// ========================================

/**
 * FUNÇÃO DE TESTE - Execute esta para verificar se está funcionando
 */
function testarPedidosAFaturar() {
  Logger.log("🧪 Iniciando teste completo OTIMIZADO...");
  Logger.log("=".repeat(50));

  // 1. Testa leitura da aba Dados1
  Logger.log("\n📋 Passo 1: Testando leitura da aba Dados1...");
  var dados = lerDados1();
  Logger.log("   Registros encontrados: " + dados.length);

  if (dados.length > 0) {
    Logger.log("   Exemplo do primeiro registro:");
    Logger.log("   - OC: " + dados[0].ordemCompra);
    Logger.log("   - Valor: " + dados[0].valor);
    Logger.log("   - Cliente: " + dados[0].cliente);
  } else {
    Logger.log("   ⚠️ PROBLEMA: Nenhum dado encontrado na aba Dados1!");
    return;
  }

  // 2. Testa criação do mapa de marcas
  Logger.log("\n🗺️ Passo 2: Testando criação do mapa OC->Marca...");
  var inicio = new Date().getTime();
  var mapaOCMarca = criarMapaOCMarca();
  var tempoMapa = (new Date().getTime() - inicio) / 1000;
  Logger.log("   Mapa criado em " + tempoMapa + " segundos");
  Logger.log("   Total de OCs no mapa: " + Object.keys(mapaOCMarca).length);

  // Testa busca de uma marca
  var ocTeste = dados[0].ordemCompra;
  Logger.log("   Testando busca para OC: " + ocTeste);
  var marca = buscarMarcaNoMapa(ocTeste, mapaOCMarca);
  Logger.log("   Marca encontrada: " + marca);

  // 3. Testa função completa
  Logger.log("\n💼 Passo 3: Testando getPedidosAFaturar()...");
  inicio = new Date().getTime();
  var resultado = getPedidosAFaturar();
  var tempoTotal = (new Date().getTime() - inicio) / 1000;

  Logger.log("   Sucesso: " + resultado.sucesso);
  Logger.log("   Timestamp: " + resultado.timestamp);
  Logger.log("   Linhas retornadas: " + resultado.dados.length);
  Logger.log("   ⏱️ Tempo de execução: " + tempoTotal + " segundos");

  if (resultado.dados.length > 0) {
    Logger.log("\n   📊 Primeiros 10 resultados:");
    resultado.dados.slice(0, 10).forEach(function(item, index) {
      Logger.log("   " + (index + 1) + ". " + item.cliente + " | " + item.marca + " | R$ " + item.valor.toFixed(2));
    });
  }

  // 4. Retorna resultado formatado em JSON
  Logger.log("\n=".repeat(50));
  Logger.log("✅ Teste concluído com sucesso!");
  Logger.log("🚀 Performance: " + tempoTotal + " segundos para " + dados.length + " registros");

  return resultado;
}

/**
 * Teste simples apenas da leitura de Dados1
 */
function testarLeituraDados1() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados1");

  if (!sheet) {
    Logger.log("❌ Aba 'Dados1' NÃO EXISTE!");
    Logger.log("Abas disponíveis na planilha:");
    SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(s) {
      Logger.log("  - " + s.getName());
    });
    return;
  }

  Logger.log("✅ Aba 'Dados1' encontrada!");
  Logger.log("Última linha: " + sheet.getLastRow());

  if (sheet.getLastRow() >= 2) {
    var dados = sheet.getRange(2, 1, Math.min(5, sheet.getLastRow() - 1), 3).getValues();
    Logger.log("\nPrimeiras " + dados.length + " linhas:");
    dados.forEach(function(row, i) {
      Logger.log("  Linha " + (i + 2) + ": OC=" + row[0] + " | Valor=" + row[1] + " | Cliente=" + row[2]);
    });
  } else {
    Logger.log("⚠️ Aba vazia (sem dados além do cabeçalho)");
  }
}

/**
 * Verifica o tamanho das abas Dados e Dados1
 */
function verificarTamanhoAbas() {
  Logger.log("📊 Verificando tamanho das abas...");
  Logger.log("=".repeat(50));

  // Verifica aba Dados
  var sheetDados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados");
  if (sheetDados) {
    var totalDados = sheetDados.getLastRow();
    Logger.log("📌 Aba DADOS:");
    Logger.log("   Total de linhas: " + totalDados);
    Logger.log("   Linhas com dados: " + (totalDados - 1));
  } else {
    Logger.log("❌ Aba 'Dados' não encontrada!");
  }

  // Verifica aba Dados1
  var sheetDados1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dados1");
  if (sheetDados1) {
    var totalDados1 = sheetDados1.getLastRow();
    Logger.log("\n📌 Aba DADOS1:");
    Logger.log("   Total de linhas: " + totalDados1);
    Logger.log("   Linhas com dados: " + (totalDados1 - 1));
  } else {
    Logger.log("\n❌ Aba 'Dados1' não encontrada!");
  }

  Logger.log("\n" + "=".repeat(50));
  Logger.log("✅ Verificação concluída!");
}
