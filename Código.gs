// --- ARQUIVO: Código.gs ---

// ========================================
// SISTEMA DE AUTENTICAÇÃO
// ========================================

/**
 * Verifica login contra a aba "senha" da planilha
 * VERSÃO SIMPLIFICADA SPA (sem tokens)
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
        status: "erro",
        mensagem: "Erro de configuração do sistema"
      };
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log("❌ Nenhum usuário cadastrado");
      return {
        status: "erro",
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

        // Retorna sucesso com nome do usuário (SEM TOKEN)
        return {
          status: "sucesso",
          nome: usuario,
          mensagem: "Login realizado com sucesso!"
        };
      }
    }

    Logger.log("❌ Credenciais inválidas para: " + usuario);
    return {
      status: "erro",
      mensagem: "Usuário ou senha incorretos"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao verificar login: " + erro.toString());
    return {
      status: "erro",
      mensagem: "Erro ao verificar credenciais: " + erro.message
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
// VERSÃO SIMPLIFICADA SPA - Sempre serve Index.html
function doGet(e) {
  Logger.log("📄 doGet chamado - Servindo Index.html (SPA)");

  // Serve sempre o Index.html - a autenticação acontece no frontend
  return HtmlService.createHtmlOutputFromFile('Index')
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
    // Lê 6 colunas: A=OC, B=Valor, C=Cliente, D=Data Recebimento, E=UNIDADE, F=QUANTIDADE
    var dados = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    var resultado = [];
    dados.forEach(function(row) {
      if (row[0] && row[1]) { // Precisa ter pelo menos OC e Valor
        resultado.push({
          ordemCompra: row[0].toString().trim(),
          valor: typeof row[1] === 'number' ? row[1] : parseFloat(row[1]) || 0,
          cliente: row[2] ? row[2].toString().trim() : "Sem Cliente",
          dataRecebimento: row[3] || null, // Coluna D (índice 3) - pode ser Date ou string
          unidade: row[4] ? row[4].toString().trim().toUpperCase() : "", // Coluna E (índice 4) - CM ou MM
          quantidade: typeof row[5] === 'number' ? row[5] : parseFloat(row[5]) || 0  // Coluna F (índice 5)
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
 * Agrupa dados da aba Dados1 por Ordem de Compra, somando valores repetidos
 * OTIMIZAÇÃO: Resolve o problema de OCs duplicadas na comparação de snapshot
 * @returns {Object} Mapa com OC como chave e {valor: total, cliente: string} como valor
 */
function agruparDados1PorOC() {
  try {
    var dados = lerDados1();
    var mapaAgrupado = {};
    var countInconsistencias = 0;

    dados.forEach(function(item) {
      var oc = item.ordemCompra;

      if (!mapaAgrupado[oc]) {
        // Primeira ocorrência desta OC
        mapaAgrupado[oc] = {
          valor: item.valor,
          cliente: item.cliente
        };
      } else {
        // OC repetida - SOMA o valor
        mapaAgrupado[oc].valor += item.valor;

        // AVISO: Detecta se a mesma OC tem clientes diferentes
        if (mapaAgrupado[oc].cliente !== item.cliente) {
          countInconsistencias++;
          Logger.log("⚠️ Aviso: OC '" + oc + "' encontrada com múltiplos clientes ('" +
                    mapaAgrupado[oc].cliente + "' e '" + item.cliente + "'). Mantendo o primeiro.");
        }
      }
    });

    Logger.log("✅ Agrupados " + Object.keys(mapaAgrupado).length + " OCs únicas de " + dados.length + " registros");
    if (countInconsistencias > 0) {
      Logger.log("⚠️ ATENÇÃO: Detectadas " + countInconsistencias + " OCs com múltiplos clientes. Verifique os dados!");
    }
    return mapaAgrupado;

  } catch (erro) {
    Logger.log("❌ Erro ao agrupar Dados1 por OC: " + erro.toString());
    return {};
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
 * Cria um mapa completo de OC com marca, pares e metros (OTIMIZADO)
 * Agrupa múltiplas linhas da mesma OC, somando pares e metros
 * @returns {Object} Mapa com OC como chave e {marca, pares, metros} como valor
 */
function criarMapaOCDadosCompleto() {
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

    var numLinhas = lastRow - 1;
    Logger.log("📥 Carregando mapa completo OC->{marca, pares, metros} de " + numLinhas + " linhas...");

    // Pega todas as 10 colunas da aba Dados
    var dados = sheet.getRange(2, 1, numLinhas, 10).getValues();

    var mapa = {};
    var contador = 0;

    dados.forEach(function(row) {
      var oc = row[9] ? row[9].toString().trim() : ""; // Coluna J (índice 9) - OC
      var marca = row[4] ? row[4].toString().trim() : "Sem Marca"; // Coluna E (índice 4) - Marca
      var qtd = typeof row[6] === 'number' ? row[6] : parseFloat(row[6]) || 0; // Coluna G (índice 6) - Qtd
      var unidade = row[7] ? row[7].toString().trim().toUpperCase() : ""; // Coluna H (índice 7) - Unidade

      if (oc && oc !== "") {
        // Se a OC ainda não existe no mapa, cria entrada
        if (!mapa[oc]) {
          mapa[oc] = {
            marca: marca,
            pares: 0,
            metros: 0
          };
        }

        // Soma nas quantidades apropriadas (permite múltiplas linhas da mesma OC)
        if (unidade.includes("PAR")) {
          mapa[oc].pares += qtd;
        } else if (unidade.includes("M") || unidade.includes("METRO")) {
          mapa[oc].metros += qtd;
        }

        contador++;
      }
    });

    Logger.log("✅ Mapa completo criado com " + Object.keys(mapa).length + " OCs únicas de " + numLinhas + " linhas processadas");
    return mapa;

  } catch (erro) {
    Logger.log("❌ Erro ao criar mapa OC completo: " + erro.toString());
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
 * CRIAR/ATUALIZAR ABA DE CONTROLE VISUAL DE FATURAMENTO
 * Mantém registro detalhado de cada OC com valores totais, faturados e saldo
 * Facilita diagnóstico e permite visualização clara de erros
 */
function criarOuAtualizarAbaControle() {
  try {
    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var nomeAba = "ControleFaturamento";
    var sheet = doc.getSheetByName(nomeAba);

    // Cria aba se não existir
    if (!sheet) {
      Logger.log("📋 Criando aba '" + nomeAba + "'...");
      sheet = doc.insertSheet(nomeAba);

      // Configura cabeçalho
      sheet.appendRow([
        "OC",
        "Cliente",
        "Marca",
        "Valor Total",
        "Valor Faturado",
        "Saldo Restante",
        "% Faturado",
        "Última Detecção",
        "Status"
      ]);

      // Formata cabeçalho
      var headerRange = sheet.getRange(1, 1, 1, 9);
      headerRange.setBackground("#1976D2");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);

      // Define larguras das colunas
      sheet.setColumnWidth(1, 120);  // OC
      sheet.setColumnWidth(2, 200);  // Cliente
      sheet.setColumnWidth(3, 150);  // Marca
      sheet.setColumnWidth(4, 120);  // Valor Total
      sheet.setColumnWidth(5, 120);  // Valor Faturado
      sheet.setColumnWidth(6, 120);  // Saldo Restante
      sheet.setColumnWidth(7, 100);  // % Faturado
      sheet.setColumnWidth(8, 150);  // Última Detecção
      sheet.setColumnWidth(9, 100);  // Status

      Logger.log("✅ Aba criada com cabeçalho");
    }

    // Sincroniza com dados atuais
    sincronizarOCsNaAbaControle(sheet);

    return {
      sucesso: true,
      mensagem: "Aba de controle atualizada"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao criar/atualizar aba controle: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * SINCRONIZAR OCs NA ABA DE CONTROLE
 * Adiciona novas OCs que apareceram e atualiza valores totais
 */
function sincronizarOCsNaAbaControle(sheet) {
  try {
    Logger.log("🔄 Sincronizando OCs na aba de controle...");

    // Lê dados atuais agrupados por OC
    var mapaAtual = agruparDados1PorOC();
    var mapaOCMarca = criarMapaOCMarca();

    // Lê o que já está na aba
    var lastRow = sheet.getLastRow();
    var dadosExistentes = {};

    if (lastRow > 1) {
      var dados = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
      dados.forEach(function(row, index) {
        var oc = row[0].toString().trim();
        dadosExistentes[oc] = {
          linha: index + 2,
          valorFaturado: typeof row[4] === 'number' ? row[4] : 0,
          ultimaDeteccao: row[7] || ""
        };
      });
    }

    var novasLinhas = [];
    var linhasAtualizadas = 0;

    // Processa cada OC atual
    Object.keys(mapaAtual).forEach(function(oc) {
      var item = mapaAtual[oc];
      var marca = buscarMarcaNoMapa(oc, mapaOCMarca);
      var valorTotal = item.valor;

      if (dadosExistentes[oc]) {
        // OC já existe - atualiza apenas valor total e saldo
        var linha = dadosExistentes[oc].linha;
        var valorFaturado = dadosExistentes[oc].valorFaturado;
        var saldoRestante = valorTotal - valorFaturado;
        var percFaturado = valorTotal > 0 ? (valorFaturado / valorTotal * 100).toFixed(1) + "%" : "0%";
        var status = saldoRestante <= 0 ? "Faturado" : (valorFaturado > 0 ? "Parcial" : "Pendente");

        sheet.getRange(linha, 4).setValue(valorTotal);  // Valor Total
        sheet.getRange(linha, 6).setValue(saldoRestante);  // Saldo Restante
        sheet.getRange(linha, 7).setValue(percFaturado);  // %
        sheet.getRange(linha, 9).setValue(status);  // Status

        linhasAtualizadas++;

      } else {
        // OC nova - adiciona
        novasLinhas.push([
          oc,
          item.cliente,
          marca,
          valorTotal,
          0,  // Valor Faturado (inicial)
          valorTotal,  // Saldo Restante
          "0%",  // % Faturado
          "",  // Última Detecção
          "Pendente"  // Status
        ]);
      }
    });

    // Adiciona novas linhas
    if (novasLinhas.length > 0) {
      sheet.getRange(lastRow + 1, 1, novasLinhas.length, 9).setValues(novasLinhas);
      Logger.log("➕ Adicionadas " + novasLinhas.length + " novas OCs");
    }

    if (linhasAtualizadas > 0) {
      Logger.log("🔄 Atualizadas " + linhasAtualizadas + " OCs existentes");
    }

    // Aplica formatação condicional
    aplicarFormatacaoCondicionalControle(sheet);

    Logger.log("✅ Sincronização concluída");

  } catch (erro) {
    Logger.log("❌ Erro ao sincronizar OCs: " + erro.toString());
  }
}

/**
 * REGISTRAR FATURAMENTO NA ABA DE CONTROLE
 * Atualiza valor faturado quando sistema detecta faturamento
 */
function registrarFaturamentoNaAbaControle(faturamentosDetectados, dataDeteccao) {
  try {
    if (!faturamentosDetectados || faturamentosDetectados.length === 0) {
      return;
    }

    Logger.log("📊 Registrando faturamento na aba de controle...");

    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = doc.getSheetByName("ControleFaturamento");

    if (!sheet) {
      Logger.log("⚠️ Aba ControleFaturamento não existe. Criando...");
      criarOuAtualizarAbaControle();
      sheet = doc.getSheetByName("ControleFaturamento");
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log("⚠️ Aba vazia. Execute criarOuAtualizarAbaControle() primeiro");
      return;
    }

    // Lê dados da aba
    var dados = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
    var mapaLinhas = {};

    dados.forEach(function(row, index) {
      var oc = row[0].toString().trim();
      mapaLinhas[oc] = {
        linha: index + 2,
        valorTotal: typeof row[3] === 'number' ? row[3] : 0,
        valorFaturado: typeof row[4] === 'number' ? row[4] : 0
      };
    });

    var linhasAtualizadas = 0;

    // Atualiza cada faturamento detectado
    faturamentosDetectados.forEach(function(item) {
      var oc = item.oc;

      if (mapaLinhas[oc]) {
        var info = mapaLinhas[oc];
        var novoValorFaturado = info.valorFaturado + item.valor;
        var saldoRestante = info.valorTotal - novoValorFaturado;
        var percFaturado = info.valorTotal > 0 ? (novoValorFaturado / info.valorTotal * 100).toFixed(1) + "%" : "0%";
        var status = saldoRestante <= 0 ? "Faturado" : (novoValorFaturado > 0 ? "Parcial" : "Pendente");

        sheet.getRange(info.linha, 5).setValue(novoValorFaturado);  // Valor Faturado
        sheet.getRange(info.linha, 6).setValue(saldoRestante);  // Saldo Restante
        sheet.getRange(info.linha, 7).setValue(percFaturado);  // %
        sheet.getRange(info.linha, 8).setValue(dataDeteccao);  // Última Detecção
        sheet.getRange(info.linha, 9).setValue(status);  // Status

        linhasAtualizadas++;
      }
    });

    Logger.log("✅ Registrados " + linhasAtualizadas + " faturamentos na aba de controle");

    // Reaplica formatação
    aplicarFormatacaoCondicionalControle(sheet);

  } catch (erro) {
    Logger.log("❌ Erro ao registrar faturamento na aba: " + erro.toString());
  }
}

/**
 * APLICAR FORMATAÇÃO CONDICIONAL À ABA DE CONTROLE
 * Destaca status com cores
 */
function aplicarFormatacaoCondicionalControle(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var dados = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

    dados.forEach(function(row, index) {
      var linha = index + 2;
      var status = row[8].toString();
      var rangeStatus = sheet.getRange(linha, 9);

      // Cores por status
      if (status === "Faturado") {
        rangeStatus.setBackground("#4CAF50").setFontColor("#FFFFFF");
      } else if (status === "Parcial") {
        rangeStatus.setBackground("#FF9800").setFontColor("#FFFFFF");
      } else if (status === "Pendente") {
        rangeStatus.setBackground("#F5F5F5").setFontColor("#000000");
      }
    });

  } catch (erro) {
    Logger.log("❌ Erro ao aplicar formatação: " + erro.toString());
  }
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
    var mapaOCDados = criarMapaOCDadosCompleto();

    // Agrupa por cliente+marca, somando valores, pares e metros
    var agrupamentoMap = {};

    dados.forEach(function(item) {
      // Busca a marca no mapa (rápido - O(1))
      var dadosOC = mapaOCDados[item.ordemCompra];
      var marca = dadosOC ? dadosOC.marca : "Sem Marca";

      var chave = item.cliente + "|" + marca;

      if (!agrupamentoMap[chave]) {
        agrupamentoMap[chave] = {
          cliente: item.cliente,
          marca: marca,
          valor: 0,
          pares: 0,
          metros: 0
        };
      }

      // Soma valores
      agrupamentoMap[chave].valor += item.valor;

      // Soma pares ou metros baseado na UNIDADE
      if (item.unidade.includes("CM")) {
        // CM = pares
        agrupamentoMap[chave].pares += item.quantidade;
      } else if (item.unidade.includes("MM")) {
        // MM = metros
        agrupamentoMap[chave].metros += item.quantidade;
      }
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
 * Retorna entradas do dia (pedidos recebidos hoje)
 * Filtra por data de recebimento = data atual
 */
function getEntradasDoDia() {
  try {
    Logger.log("📦 Iniciando getEntradasDoDia...");

    var dados = lerDados1();

    if (dados.length === 0) {
      return {
        sucesso: true,
        timestamp: obterTimestamp(),
        dados: []
      };
    }

    // Obtém a data de hoje (sem hora) para comparação
    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    Logger.log("📅 Data de hoje: " + Utilities.formatDate(hoje, Session.getScriptTimeZone(), "dd/MM/yyyy"));

    // Carrega mapa de marcas da aba Dados
    var mapaOCDados = criarMapaOCDadosCompleto();

    // Filtra apenas pedidos recebidos hoje
    var resultado = [];
    dados.forEach(function(item) {
      if (item.dataRecebimento) {
        // Converte data de recebimento para Date (se for string) e normaliza
        var dataReceb;
        if (item.dataRecebimento instanceof Date) {
          dataReceb = new Date(item.dataRecebimento);
        } else {
          // Tenta converter string DD/MM/YYYY para Date
          var partes = item.dataRecebimento.toString().split('/');
          if (partes.length === 3) {
            dataReceb = new Date(partes[2], partes[1] - 1, partes[0]);
          }
        }

        if (dataReceb) {
          dataReceb.setHours(0, 0, 0, 0);

          // Compara se é hoje
          if (dataReceb.getTime() === hoje.getTime()) {
            // Busca a marca
            var dadosOC = mapaOCDados[item.ordemCompra];
            var marca = dadosOC ? dadosOC.marca : "Sem Marca";

            resultado.push({
              cliente: item.cliente,
              marca: marca,
              ordemCompra: item.ordemCompra,
              valor: item.valor,
              dataRecebimento: Utilities.formatDate(dataReceb, Session.getScriptTimeZone(), "dd/MM/yyyy")
            });
          }
        }
      }
    });

    // Ordena por cliente (alfabético)
    resultado.sort(function(a, b) {
      return a.cliente.localeCompare(b.cliente);
    });

    Logger.log("✅ getEntradasDoDia concluído: " + resultado.length + " entradas hoje");

    return {
      sucesso: true,
      timestamp: obterTimestamp(),
      dados: resultado
    };

  } catch (erro) {
    Logger.log("❌ Erro em getEntradasDoDia: " + erro.toString());
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

    // CORREÇÃO: Verifica e limpa faturamento de dias anteriores
    // Isso previne que o card exiba dados antigos como se fossem de hoje
    var dataAtual = new Date();
    var diaAtual = ("0" + dataAtual.getDate()).slice(-2) + "/" +
                   ("0" + (dataAtual.getMonth() + 1)).slice(-2) + "/" +
                   dataAtual.getFullYear();

    var diaArmazenado = props.getProperty('FATURAMENTO_DATA');

    if (diaArmazenado && diaArmazenado !== diaAtual) {
      Logger.log("📅 Detectado mudança de dia: " + diaArmazenado + " → " + diaAtual);
      Logger.log("🔄 Limpando faturamento acumulado do dia anterior...");
      props.deleteProperty('ULTIMO_FATURAMENTO');
      props.deleteProperty('ULTIMO_FATURAMENTO_TIMESTAMP');
      props.deleteProperty('FATURAMENTO_DATA');
      Logger.log("✅ Faturamento acumulado limpo");
    }

    // SINCRONIZAÇÃO AUTOMÁTICA: Atualiza aba de controle com novas OCs
    // Isso garante que pedidos novos apareçam automaticamente na aba
    try {
      var doc = SpreadsheetApp.getActiveSpreadsheet();
      var sheetControle = doc.getSheetByName("ControleFaturamento");

      if (sheetControle) {
        Logger.log("🔄 Sincronizando aba de controle com novos pedidos...");
        sincronizarOCsNaAbaControle(sheetControle);
      } else {
        Logger.log("ℹ️ Aba ControleFaturamento não existe. Execute criarOuAtualizarAbaControle() para criar.");
      }
    } catch (erroSinc) {
      Logger.log("⚠️ Erro ao sincronizar aba de controle: " + erroSinc.toString());
      // Continua execução mesmo se sincronização falhar
    }

    var snapshotAnterior = props.getProperty('SNAPSHOT_DADOS1');
    var timestampAnterior = props.getProperty('SNAPSHOT_TIMESTAMP');

    // Lê estado atual AGRUPADO por OC (soma valores repetidos)
    // OTIMIZAÇÃO: Resolve problema de OCs duplicadas
    var mapaAtual = agruparDados1PorOC();

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

    // NOVA LÓGICA: Compara totais AGRUPADOS por OC
    // Antes: Comparava linha por linha (OCs duplicadas sobrescreviam)
    // Agora: Compara soma total de cada OC (valores repetidos são somados)
    // Benefício: Detecção precisa mesmo com múltiplas linhas da mesma OC
    Object.keys(mapaAnterior).forEach(function(oc) {
      var itemAnterior = mapaAnterior[oc];
      var itemAtual = mapaAtual[oc];

      var valorFaturado = 0;

      if (!itemAtual) {
        // OC sumiu completamente = faturou tudo
        valorFaturado = itemAnterior.valor;
      } else if (itemAtual.valor < itemAnterior.valor) {
        // Valor total diminuiu = faturou a diferença
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

      // NOVO: Registra faturamento na aba de controle visual (com OCs individuais)
      // Usa a lista 'faturado' que contém os OCs antes do agrupamento
      registrarFaturamentoNaAbaControle(faturado, diaAtual + " " + obterTimestamp().split(" às ")[1]);
    } else if (faturamentoAcumulado.length > 0) {
      // Não houve novo faturamento, mas há acumulado do dia
      Logger.log("ℹ️ Nenhum novo faturamento nesta verificação, mantendo acumulado do dia");
      resultado = faturamentoAcumulado;
    }

    Logger.log("✅ getFaturamentoDia concluído: " + resultado.length + " itens calculados");

    // IMPORTANTE: Lê os dados REAIS do histórico (incluindo edições manuais)
    // Não retorna o calculado, mas sim o que está efetivamente salvo
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");
    if (sheet && sheet.getLastRow() > 1) {
      var historicoDados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      var dadosDodia = [];

      historicoDados.forEach(function(row) {
        // Normaliza data
        var dataRegistro = row[0];
        if (dataRegistro instanceof Date) {
          var d = dataRegistro;
          var dia = ("0" + d.getDate()).slice(-2);
          var mes = ("0" + (d.getMonth() + 1)).slice(-2);
          var ano = d.getFullYear();
          dataRegistro = dia + "/" + mes + "/" + ano;
        } else {
          dataRegistro = dataRegistro.toString().trim();
        }

        // Se é o dia de hoje
        if (dataRegistro === diaAtual) {
          dadosDodia.push({
            cliente: row[1].toString(),
            marca: row[2].toString(),
            valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0
          });
        }
      });

      if (dadosDodia.length > 0) {
        Logger.log("📊 Retornando dados do histórico (incluindo edições manuais): " + dadosDodia.length + " itens");
        resultado = dadosDodia;
      }
    }

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
 * ATUALIZADO: Agora lê do HISTÓRICO (inclui edições manuais)
 */
function getUltimoFaturamento() {
  try {
    Logger.log("📊 getUltimoFaturamento: Lendo dados do histórico...");

    // Data de hoje
    var dataAtual = new Date();
    var diaAtual = ("0" + dataAtual.getDate()).slice(-2) + "/" +
                   ("0" + (dataAtual.getMonth() + 1)).slice(-2) + "/" +
                   dataAtual.getFullYear();

    // Lê dados REAIS do histórico (incluindo edições manuais)
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");

    if (!sheet || sheet.getLastRow() < 2) {
      Logger.log("⚠️ Histórico vazio ou não encontrado");
      return {
        sucesso: true,
        timestamp: null,
        dados: [],
        mensagem: "Nenhum faturamento detectado ainda. Aguardando primeira verificação."
      };
    }

    // Lê todos os dados do histórico
    var historicoDados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    var dadosDodia = [];
    var ultimaDataComDados = null;
    var timestampUltimoRegistro = null;

    // Primeiro, tenta buscar dados do dia atual
    historicoDados.forEach(function(row) {
      // Normaliza data
      var dataRegistro = row[0];
      if (dataRegistro instanceof Date) {
        var d = dataRegistro;
        var dia = ("0" + d.getDate()).slice(-2);
        var mes = ("0" + (d.getMonth() + 1)).slice(-2);
        var ano = d.getFullYear();
        dataRegistro = dia + "/" + mes + "/" + ano;
      } else {
        dataRegistro = dataRegistro.toString().trim();
      }

      // Se é o dia de hoje
      if (dataRegistro === diaAtual) {
        dadosDodia.push({
          cliente: row[1].toString(),
          marca: row[2].toString(),
          valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0,
          data: dataRegistro
        });
        ultimaDataComDados = dataRegistro;
        // Pega o timestamp da coluna F (índice 5)
        if (row[5]) {
          timestampUltimoRegistro = row[5].toString();
        }
      }
    });

    // Se não houver dados de hoje, busca os dados do último dia registrado
    if (dadosDodia.length === 0) {
      Logger.log("ℹ️ Sem dados de hoje, buscando último faturamento registrado...");

      // Agrupa dados por data para encontrar a data mais recente
      var dadosPorData = {};

      historicoDados.forEach(function(row) {
        var dataRegistro = row[0];
        if (dataRegistro instanceof Date) {
          var d = dataRegistro;
          var dia = ("0" + d.getDate()).slice(-2);
          var mes = ("0" + (d.getMonth() + 1)).slice(-2);
          var ano = d.getFullYear();
          dataRegistro = dia + "/" + mes + "/" + ano;
        } else {
          dataRegistro = dataRegistro.toString().trim();
        }

        if (!dadosPorData[dataRegistro]) {
          dadosPorData[dataRegistro] = [];
        }

        dadosPorData[dataRegistro].push({
          cliente: row[1].toString(),
          marca: row[2].toString(),
          valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0,
          data: dataRegistro,
          timestamp: row[5] ? row[5].toString() : null
        });
      });

      // Encontra a data mais recente (converte para Date para comparar)
      var datasOrdenadas = Object.keys(dadosPorData).sort(function(a, b) {
        var partesA = a.split('/');
        var partesB = b.split('/');
        var dateA = new Date(partesA[2], partesA[1] - 1, partesA[0]);
        var dateB = new Date(partesB[2], partesB[1] - 1, partesB[0]);
        return dateB - dateA; // Mais recente primeiro
      });

      if (datasOrdenadas.length > 0) {
        ultimaDataComDados = datasOrdenadas[0];
        dadosDodia = dadosPorData[ultimaDataComDados];

        // Pega o timestamp do último registro dessa data
        var ultimoRegistro = dadosDodia[dadosDodia.length - 1];
        if (ultimoRegistro.timestamp) {
          timestampUltimoRegistro = ultimoRegistro.timestamp;
        }

        Logger.log("📅 Exibindo dados do último faturamento: " + ultimaDataComDados + " (" + dadosDodia.length + " registros)");
      }
    }

    Logger.log("✅ getUltimoFaturamento retornou " + dadosDodia.length + " registros");

    if (dadosDodia.length === 0) {
      return {
        sucesso: true,
        timestamp: null,
        dados: [],
        mensagem: "Nenhum faturamento registrado no histórico."
      };
    }

    // Formata o timestamp para exibição
    var ehHoje = ultimaDataComDados === diaAtual;
    var timestampExibicao;

    if (ehHoje) {
      // É de hoje - mostra timestamp ou "hoje"
      if (timestampUltimoRegistro) {
        timestampExibicao = "Faturamento de hoje: " + timestampUltimoRegistro;
      } else {
        timestampExibicao = "Faturamento de hoje";
      }
    } else {
      // É histórico - mostra a data
      timestampExibicao = "Faturamento de " + ultimaDataComDados;
    }

    return {
      sucesso: true,
      timestamp: timestampExibicao,
      dados: dadosDodia,
      ehHoje: ehHoje,
      dataExibida: ultimaDataComDados
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
      // Adiciona cabeçalho (com coluna Observação)
      sheet.appendRow(["Data", "Cliente", "Marca", "Valor Faturado", "Observação", "Timestamp"]);
      // Formata cabeçalho
      var headerRange = sheet.getRange(1, 1, 1, 6);
      headerRange.setBackground("#d32f2f");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    var timestamp = obterTimestamp();
    var novasLinhas = [];

    // Verifica registros já existentes para esta data
    var lastRow = sheet.getLastRow();
    var registrosExistentes = {};

    if (lastRow > 1) {
      // Lê todos os registros do histórico
      var dadosExistentes = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

      dadosExistentes.forEach(function(row) {
        // Normaliza data
        var dataRegistro = row[0];
        if (dataRegistro instanceof Date) {
          var d = dataRegistro;
          var dia = ("0" + d.getDate()).slice(-2);
          var mes = ("0" + (d.getMonth() + 1)).slice(-2);
          var ano = d.getFullYear();
          dataRegistro = dia + "/" + mes + "/" + ano;
        } else {
          dataRegistro = dataRegistro.toString().trim();
        }

        // Se é o mesmo dia que estamos salvando
        if (dataRegistro === data) {
          var chave = row[1].toString().toUpperCase() + "|" + row[2].toString().toUpperCase();
          registrosExistentes[chave] = {
            valor: row[3],
            observacao: row[4] ? row[4].toString() : ""
          };
        }
      });

      Logger.log("📋 Encontrados " + Object.keys(registrosExistentes).length + " registros existentes para " + data);
    }

    // Processa novos dados
    dados.forEach(function(item) {
      var chave = item.cliente.toUpperCase() + "|" + item.marca.toUpperCase();

      // Se já existe no histórico
      if (registrosExistentes[chave]) {
        var registroExistente = registrosExistentes[chave];

        // Se tem observação = foi editado manualmente = NÃO sobrescreve
        if (registroExistente.observacao && registroExistente.observacao.trim() !== "") {
          Logger.log("✏️ Mantendo valor editado manualmente: " + item.cliente + " | " + item.marca + " = R$ " + registroExistente.valor);
          // Não adiciona à lista de novas linhas (mantém o existente)
        } else {
          // Sem observação = valor automático = pode atualizar
          Logger.log("🔄 Atualizando valor automático: " + item.cliente + " | " + item.marca + " = R$ " + item.valor);
          // Remove o antigo (será adicionado novamente com novo valor)
          registrosExistentes[chave] = null;

          novasLinhas.push([
            data,
            item.cliente,
            item.marca,
            item.valor,
            "", // Observação vazia (automático)
            timestamp
          ]);
        }
      } else {
        // Registro novo - adiciona
        Logger.log("➕ Adicionando novo registro: " + item.cliente + " | " + item.marca + " = R$ " + item.valor);
        novasLinhas.push([
          data,
          item.cliente,
          item.marca,
          item.valor,
          "", // Observação vazia (automático)
          timestamp
        ]);
      }
    });

    // Remove registros automáticos antigos que serão atualizados
    if (lastRow > 1) {
      for (var i = lastRow; i >= 2; i--) {
        var row = sheet.getRange(i, 1, 1, 6).getValues()[0];

        // Normaliza data
        var dataLinha = row[0];
        if (dataLinha instanceof Date) {
          var d = dataLinha;
          var dia = ("0" + d.getDate()).slice(-2);
          var mes = ("0" + (d.getMonth() + 1)).slice(-2);
          var ano = d.getFullYear();
          dataLinha = dia + "/" + mes + "/" + ano;
        } else {
          dataLinha = dataLinha.toString().trim();
        }

        // Se é o mesmo dia e NÃO tem observação (automático)
        if (dataLinha === data) {
          var obs = row[4] ? row[4].toString().trim() : "";
          if (!obs || obs === "") {
            Logger.log("🗑️ Removendo registro automático antigo linha " + i);
            sheet.deleteRow(i);
          }
        }
      }
    }

    // Adiciona as novas linhas
    if (novasLinhas.length > 0) {
      var ultimaLinha = sheet.getLastRow();
      sheet.getRange(ultimaLinha + 1, 1, novasLinhas.length, 6).setValues(novasLinhas);

      // Formata valores como moeda
      var valorRange = sheet.getRange(ultimaLinha + 1, 4, novasLinhas.length, 1);
      valorRange.setNumberFormat("R$ #,##0.00");

      Logger.log("✅ Salvou " + novasLinhas.length + " linhas no histórico para " + data);
    } else {
      Logger.log("ℹ️ Nenhum registro novo para adicionar (todos já existem ou foram editados manualmente)");
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

    // Lê todos os dados (pula cabeçalho) - agora com 6 colunas incluindo Observação
    var dados = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

    var historico = [];

    dados.forEach(function(row) {
      // Formata timestamp se vier como Date object
      var timestampFormatado = row[5];
      if (row[5] instanceof Date) {
        var d = row[5];
        var dia = ("0" + d.getDate()).slice(-2);
        var mes = ("0" + (d.getMonth() + 1)).slice(-2);
        var ano = d.getFullYear();
        var hora = ("0" + d.getHours()).slice(-2);
        var minuto = ("0" + d.getMinutes()).slice(-2);
        timestampFormatado = dia + "/" + mes + "/" + ano + " às " + hora + ":" + minuto;
      } else {
        timestampFormatado = row[5] ? row[5].toString() : "";
      }

      // Formata data se vier como Date object
      var dataFormatada = row[0];
      if (row[0] instanceof Date) {
        var d = row[0];
        var dia = ("0" + d.getDate()).slice(-2);
        var mes = ("0" + (d.getMonth() + 1)).slice(-2);
        var ano = d.getFullYear();
        dataFormatada = dia + "/" + mes + "/" + ano;
      } else {
        dataFormatada = row[0] ? row[0].toString() : "";
      }

      historico.push({
        data: dataFormatada,
        cliente: row[1].toString(),
        marca: row[2].toString(),
        valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0,
        observacao: row[4] ? row[4].toString() : "",
        timestamp: timestampFormatado
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
 * Configura triggers automáticos (a cada 1 hora)
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

  // Cria trigger para executar A CADA 1 HORA
  ScriptApp.newTrigger('executarVerificacaoFaturamento')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("✅ Triggers configurados com sucesso!");
  Logger.log("⏰ Verificações automáticas A CADA 1 HORA (24x por dia)");
  Logger.log("ℹ️  Sistema detectará faturamento muito mais rápido!");
}

/**
 * Configura triggers para 2x ao dia (8h e 19h) - MODO ECONÔMICO
 * Use esta função se quiser menos verificações (economiza quotas do Google)
 */
function setupTriggers2xDia() {
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
  Logger.log("⏰ Verificações automáticas às 8h e 19h (modo econômico)");
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

/**
 * ========================================
 * FUNÇÕES PARA EDIÇÃO MANUAL DE FATURAMENTO
 * ========================================
 */

/**
 * Edita um registro específico de faturamento
 * @param {string} data - Data do registro (DD/MM/AAAA)
 * @param {string} cliente - Nome do cliente
 * @param {string} marca - Marca
 * @param {number} novoValor - Novo valor corrigido
 * @param {string} observacao - Observação sobre o ajuste
 * @returns {Object} Resultado da operação
 */
function editarRegistroFaturamento(data, cliente, marca, novoValor, observacao) {
  try {
    Logger.log("✏️ Editando registro: " + data + " | " + cliente + " | " + marca);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");

    if (!sheet) {
      return {
        sucesso: false,
        mensagem: "Aba 'HistoricoFaturamento' não encontrada"
      };
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        sucesso: false,
        mensagem: "Nenhum registro encontrado no histórico"
      };
    }

    // Busca o registro
    var dados = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var registroEncontrado = false;
    var linhaParaEditar = -1;

    // Normaliza os valores de busca
    var dataBusca = data.trim();
    var clienteBusca = cliente.trim().toUpperCase();
    var marcaBusca = marca.trim().toUpperCase();

    Logger.log("🔍 Buscando: Data=" + dataBusca + " | Cliente=" + clienteBusca + " | Marca=" + marcaBusca);

    for (var i = 0; i < dados.length; i++) {
      // Normaliza a data da planilha
      var dataPlanilha = dados[i][0];
      if (dataPlanilha instanceof Date) {
        var d = dataPlanilha;
        var dia = ("0" + d.getDate()).slice(-2);
        var mes = ("0" + (d.getMonth() + 1)).slice(-2);
        var ano = d.getFullYear();
        dataPlanilha = dia + "/" + mes + "/" + ano;
      } else {
        dataPlanilha = dataPlanilha.toString().trim();
      }

      var clientePlanilha = dados[i][1] ? dados[i][1].toString().trim().toUpperCase() : "";
      var marcaPlanilha = dados[i][2] ? dados[i][2].toString().trim().toUpperCase() : "";

      Logger.log("📋 Linha " + (i+2) + ": Data=" + dataPlanilha + " | Cliente=" + clientePlanilha + " | Marca=" + marcaPlanilha);

      if (dataPlanilha === dataBusca &&
          clientePlanilha === clienteBusca &&
          marcaPlanilha === marcaBusca) {
        linhaParaEditar = i + 2; // +2 porque array começa em 0 e pula cabeçalho
        registroEncontrado = true;
        Logger.log("✅ Registro encontrado na linha " + linhaParaEditar);
        break;
      }
    }

    if (!registroEncontrado) {
      Logger.log("❌ Registro NÃO encontrado após buscar " + dados.length + " linhas");
      return {
        sucesso: false,
        mensagem: "Registro não encontrado. Data: " + dataBusca + ", Cliente: " + clienteBusca + ", Marca: " + marcaBusca
      };
    }

    // Atualiza o valor e observação
    sheet.getRange(linhaParaEditar, 4).setValue(novoValor); // Coluna D: Valor
    sheet.getRange(linhaParaEditar, 5).setValue(observacao); // Coluna E: Observação

    // Formata valor como moeda
    sheet.getRange(linhaParaEditar, 4).setNumberFormat("R$ #,##0.00");

    Logger.log("✅ Registro editado com sucesso!");

    return {
      sucesso: true,
      mensagem: "Registro atualizado com sucesso!"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao editar registro: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro ao editar: " + erro.message
    };
  }
}

/**
 * Deleta um registro específico de faturamento
 * @param {string} data - Data do registro (DD/MM/AAAA)
 * @param {string} cliente - Nome do cliente
 * @param {string} marca - Marca
 * @returns {Object} Resultado da operação
 */
function deletarRegistroFaturamento(data, cliente, marca) {
  try {
    Logger.log("🗑️ Deletando registro: " + data + " | " + cliente + " | " + marca);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");

    if (!sheet) {
      return {
        sucesso: false,
        mensagem: "Aba 'HistoricoFaturamento' não encontrada"
      };
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return {
        sucesso: false,
        mensagem: "Nenhum registro encontrado no histórico"
      };
    }

    // Busca o registro
    var dados = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    var linhaParaDeletar = -1;

    // Normaliza os valores de busca
    var dataBusca = data.trim();
    var clienteBusca = cliente.trim().toUpperCase();
    var marcaBusca = marca.trim().toUpperCase();

    Logger.log("🔍 Buscando para deletar: Data=" + dataBusca + " | Cliente=" + clienteBusca + " | Marca=" + marcaBusca);

    for (var i = 0; i < dados.length; i++) {
      // Normaliza a data da planilha
      var dataPlanilha = dados[i][0];
      if (dataPlanilha instanceof Date) {
        var d = dataPlanilha;
        var dia = ("0" + d.getDate()).slice(-2);
        var mes = ("0" + (d.getMonth() + 1)).slice(-2);
        var ano = d.getFullYear();
        dataPlanilha = dia + "/" + mes + "/" + ano;
      } else {
        dataPlanilha = dataPlanilha.toString().trim();
      }

      var clientePlanilha = dados[i][1] ? dados[i][1].toString().trim().toUpperCase() : "";
      var marcaPlanilha = dados[i][2] ? dados[i][2].toString().trim().toUpperCase() : "";

      if (dataPlanilha === dataBusca &&
          clientePlanilha === clienteBusca &&
          marcaPlanilha === marcaBusca) {
        linhaParaDeletar = i + 2; // +2 porque array começa em 0 e pula cabeçalho
        Logger.log("✅ Registro encontrado na linha " + linhaParaDeletar);
        break;
      }
    }

    if (linhaParaDeletar === -1) {
      Logger.log("❌ Registro NÃO encontrado após buscar " + dados.length + " linhas");
      return {
        sucesso: false,
        mensagem: "Registro não encontrado. Data: " + dataBusca + ", Cliente: " + clienteBusca + ", Marca: " + marcaBusca
      };
    }

    // Deleta a linha
    sheet.deleteRow(linhaParaDeletar);

    Logger.log("✅ Registro deletado com sucesso!");

    return {
      sucesso: true,
      mensagem: "Registro deletado com sucesso!"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao deletar registro: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro ao deletar: " + erro.message
    };
  }
}
