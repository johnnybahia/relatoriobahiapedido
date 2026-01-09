/**
 * Script para gerar Relatório de Produção por Ordem de Compra (OC).
 * Desenvolvido para Google Sheets.
 */

// --- CONFIGURAÇÃO DAS COLUNAS (AJUSTE SE NECESSÁRIO) ---
const COLUNAS = {
  ORDEM_COMPRA: "ORD. COMPRA", 
  CLIENTE: "CLIENTE",
  PEDIDO: "PEDIDO",
  COD_CLIENTE: "CÓD. CLIENTE",
  COD_MARFIM: "CÓD. MARFIM", 
  DESCRICAO: "DESCRIÇÃO",
  TAMANHO: "TAMANHO",
  QTD_ABERTA: "QTD. ABERTA", // Ajustado para incluir o ponto se estiver assim no cabeçalho
  LOTES: "LOTES",
  CODIGO_OS: "código OS",
  DT_RECEBIMENTO: "DT. RECEBIMENTO",
  DT_ENTREGA: "DT. ENTREGA",
  PRAZO: "PRAZO"
};

// Nome da aba onde estão as MARCAS
const ABA_MARCAS_NOME = "MARCAS";

/**
 * ESTA FUNÇÃO CRIA O MENU AUTOMATICAMENTE
 * Ela roda toda vez que você abre a planilha.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏭 Relatórios')
      .addItem('🖨️ Imprimir Relatório por OC', 'gerarRelatorioOC') 
      .addToUi();
}

function gerarRelatorioOC() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 1. Pedir a OC ao usuário
  const result = ui.prompt(
      'Imprimir Relatório',
      'Digite o número da Ordem de Compra (OC):',
      ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  const ocDigitada = result.getResponseText().trim();
  if (ocDigitada === "") {
    ui.alert("Por favor, digite uma OC válida.");
    return;
  }

  // 2. Buscar a Marca na aba MARCAS
  let marcaEncontrada = "N/A";
  const sheetMarcas = ss.getSheetByName(ABA_MARCAS_NOME);
  
  if (sheetMarcas) {
    const dadosMarcas = sheetMarcas.getDataRange().getValues();
    
    let colIndexOcMarca = -1;
    let colIndexNomeMarca = -1;
    
    // Procura cabeçalhos na linha 1 da aba MARCAS
    const headerMarcas = dadosMarcas[0].map(c => String(c).toUpperCase().trim());
    colIndexOcMarca = headerMarcas.indexOf("ORDEM DE COMPRA"); 
    if (colIndexOcMarca === -1) colIndexOcMarca = headerMarcas.indexOf("ORD. COMPRA"); 
    
    colIndexNomeMarca = headerMarcas.indexOf("MARCA");

    // Se não achar cabeçalho, assume colunas padrão A=0 e B=1
    if (colIndexOcMarca === -1) colIndexOcMarca = 0; 
    if (colIndexNomeMarca === -1) colIndexNomeMarca = 1; 

    for (let i = 1; i < dadosMarcas.length; i++) {
      if (String(dadosMarcas[i][colIndexOcMarca]).trim() == ocDigitada) {
        marcaEncontrada = dadosMarcas[i][colIndexNomeMarca];
        break;
      }
    }
  } else {
    ui.alert(`Aba '${ABA_MARCAS_NOME}' não encontrada. A Marca ficará em branco.`);
  }

  // 3. Pegar dados da Planilha Ativa
  const dados = sheet.getDataRange().getDisplayValues(); 
  
  if (dados.length < 3) {
    ui.alert("A planilha não tem linhas suficientes para ter o cabeçalho na linha 3.");
    return;
  }

  // LINHA 3 = ÍNDICE 2
  const INDICE_CABECALHO = 2; 
  const cabecalho = dados[INDICE_CABECALHO].map(c => String(c).trim().toUpperCase());
  
  // Mapear índices das colunas
  const mapa = {};
  for (let key in COLUNAS) {
    let index = cabecalho.indexOf(COLUNAS[key].toUpperCase());
    if (index === -1) {
        // Fallback: Tenta encontrar coluna que CONTENHA o texto (caso tenha quebras de linha ou espaços extras)
        index = cabecalho.findIndex(c => c.includes(COLUNAS[key].toUpperCase()));
    }
    mapa[key] = index;
  }

  // Depuração rápida: se não achar QTD ABERTA, avisa qual índice encontrou
  if (mapa.QTD_ABERTA === -1) {
     // Tenta procurar sem o ponto como última tentativa
     let indexSemPonto = cabecalho.indexOf("QTD ABERTA");
     if (indexSemPonto === -1) indexSemPonto = cabecalho.findIndex(c => c.includes("QTD") && c.includes("ABERTA"));
     
     if (indexSemPonto > -1) {
       mapa.QTD_ABERTA = indexSemPonto;
     } else {
       ui.alert(`Atenção: Coluna 'QTD. ABERTA' não foi encontrada. Verifique se o nome está exato na linha 3.`);
     }
  }

  if (mapa.ORDEM_COMPRA === -1) {
    ui.alert(`Coluna '${COLUNAS.ORDEM_COMPRA}' não encontrada na linha 3.`);
    return;
  }

  // 4. Filtrar Linhas
  const itensRelatorio = [];
  let clienteNome = "";

  for (let i = INDICE_CABECALHO + 1; i < dados.length; i++) {
    const linha = dados[i];
    if (linha.length <= mapa.ORDEM_COMPRA) continue;

    const ocLinha = String(linha[mapa.ORDEM_COMPRA]).trim();

    if (ocLinha === ocDigitada) {
      if (clienteNome === "" && mapa.CLIENTE !== -1) {
        clienteNome = linha[mapa.CLIENTE];
      }

      itensRelatorio.push({
        pedido: mapa.PEDIDO > -1 ? linha[mapa.PEDIDO] : "",
        codCliente: mapa.COD_CLIENTE > -1 ? linha[mapa.COD_CLIENTE] : "",
        codMarfim: mapa.COD_MARFIM > -1 ? linha[mapa.COD_MARFIM] : "",
        descricao: mapa.DESCRICAO > -1 ? linha[mapa.DESCRICAO] : "",
        tamanho: mapa.TAMANHO > -1 ? linha[mapa.TAMANHO] : "",
        qtdAberta: mapa.QTD_ABERTA > -1 ? linha[mapa.QTD_ABERTA] : "",
        lotes: mapa.LOTES > -1 ? linha[mapa.LOTES] : "",
        codOs: mapa.CODIGO_OS > -1 ? linha[mapa.CODIGO_OS] : "",
        dtRec: mapa.DT_RECEBIMENTO > -1 ? linha[mapa.DT_RECEBIMENTO] : "",
        dtEnt: mapa.DT_ENTREGA > -1 ? linha[mapa.DT_ENTREGA] : "",
        prazo: mapa.PRAZO > -1 ? linha[mapa.PRAZO] : ""
      });
    }
  }

  if (itensRelatorio.length === 0) {
    ui.alert(`Nenhum item encontrado para a OC: ${ocDigitada}`);
    return;
  }

  // 5. Gerar HTML do Relatório
  const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");
  
  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; font-size: 12px; }
        .header { margin-bottom: 20px; border-bottom: 2px solid #000; padding-bottom: 10px; }
        .header h2 { margin: 0; }
        .info-grid { display: flex; justify-content: space-between; margin-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { border: 1px solid #333; padding: 5px; text-align: center; font-size: 11px; }
        th { background-color: #f0f0f0; font-weight: bold; }
        .text-left { text-align: left; }
        .destaque-marca { font-size: 14px; margin-top: 5px; color: #333; }
        .btn-print { 
            padding: 10px 20px; background: #007bff; color: white; border: none; cursor: pointer; font-size: 14px; 
            margin-bottom: 20px; display: block;
        }
        @media print {
            .btn-print { display: none; }
            @page { size: landscape; margin: 1cm; }
        }
      </style>
    </head>
    <body>
      <button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR AGORA</button>
      
      <div class="header">
        <div class="info-grid">
           <div><strong>RELATÓRIO DE PRODUÇÃO</strong></div>
           <div>Data: ${dataHoje}</div>
        </div>
        <div class="info-grid">
           <div><strong>CLIENTE:</strong> ${clienteNome}</div>
           <div><strong>ORD. COMPRA (OC):</strong> ${ocDigitada}</div>
        </div>
        <div class="destaque-marca">
           <strong>MARCA: ${marcaEncontrada.toUpperCase()}</strong>
        </div>
      </div>

      <table>
        <thead>
          <tr>
            <th>PEDIDO</th>
            <th>CÓD. CLIENTE</th>
            <th>CÓD. MARFIM</th>
            <th>DESCRIÇÃO</th>
            <th>TAMANHO</th>
            <th>QTD. ABERTA</th>
            <th>LOTES</th>
            <th>CÓD. OS</th>
            <th>DT. REC.</th>
            <th>DT. ENT.</th>
            <th>PRAZO</th>
          </tr>
        </thead>
        <tbody>
  `;

  itensRelatorio.forEach(item => {
    html += `
      <tr>
        <td>${item.pedido}</td>
        <td>${item.codCliente}</td>
        <td>${item.codMarfim}</td>
        <td class="text-left">${item.descricao}</td>
        <td>${item.tamanho}</td>
        <td>${item.qtdAberta}</td>
        <td>${item.lotes}</td>
        <td>${item.codOs}</td>
        <td>${item.dtRec}</td>
        <td>${item.dtEnt}</td>
        <td>${item.prazo}</td>
      </tr>
    `;
  });

  html += `
        </tbody>
      </table>
    </body>
    </html>
  `;

  const output = HtmlService.createHtmlOutput(html)
      .setWidth(1000)
      .setHeight(600);
  ui.showModalDialog(output, 'Visualização de Impressão');
}
function monitorarIMPORTRANGE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetBanco = ss.getSheetByName("BANCO DE DADOS");
  var sheetRelatorio = ss.getSheetByName("RELATÓRIO GERAL DA PRODUÇÃO");

  if (!sheetBanco || !sheetRelatorio) {
    Logger.log("Uma ou ambas as abas não foram encontradas.");
    return;
  }

  var range = sheetBanco.getRange("A1:n5000");
  var valoresAtuais = range.getValues();

  var cache = PropertiesService.getScriptProperties();
  var valoresAntigos = cache.getProperty("dados_antigos");

  if (valoresAntigos) {
    valoresAntigos = JSON.parse(valoresAntigos);

    // Convertendo arrays para strings para evitar problemas de formatação
    var atualStr = JSON.stringify(valoresAtuais);
    var antigoStr = JSON.stringify(valoresAntigos);

    if (atualStr !== antigoStr) {
      var now = Utilities.formatDate(new Date(), "America/Fortaleza", "dd/MM/yyyy HH:mm:ss");
      sheetRelatorio.getRange("H2").setValue(now); // Atualiza horário

      cache.setProperty("dados_antigos", JSON.stringify(valoresAtuais)); // Salva novo estado
      Logger.log("Alteração detectada! Horário atualizado.");
    } else {
      Logger.log("Nenhuma alteração detectada. H2 permanece o mesmo.");
    }
  } else {
    // Caso seja a primeira vez rodando, salva o estado inicial
    cache.setProperty("dados_antigos", JSON.stringify(valoresAtuais));
    Logger.log("Primeira execução: cache inicializado.");
  }
}
/****************************************************
 * Painel de Pedidos Marfim Bahia – Web App (Itens)
 * Autor: Johnny
 * Versão: 3.1.0
 * Atualizado: 2025-10-24
 *
 * MUDANÇAS:
 * - Renomeado para "Bahia"
 * - Limpeza do nome da versão
 ****************************************************/

// ====== CONFIG ======
const SPREADSHEET_ID = '1YoSxArGafauFK8DNf6w9C3CfKGvJ4ECKdMUs2n6Zh58';
const SHEET_NAME     = 'RELATÓRIO GERAL DA PRODUÇÃO1';
const HEADER_ROW     = 3;
const TZ             = 'America/Fortaleza';

// --- MUDANÇA 2: VERSÃO ATUALIZADA ---
const APP_VERSION    = '3.1.0';

// ====== BOOTSTRAP HTML ======
function doGet(e) {
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.timezone   = TZ;
  tpl.appVersion = APP_VERSION;
  
  return tpl.evaluate()
    // --- MUDANÇA 1: TÍTULO ATUALIZADO ---
    .setTitle('Painel de Pedidos Marfim Bahia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fn) { 
  return HtmlService.createHtmlOutputFromFile(fn).getContent(); 
}

// ====== HELPERS (Simplificados) ======
function _openSS_() {
  try {
    if (SPREADSHEET_ID && !/^COLE_AQUI/.test(SPREADSHEET_ID)) {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
  } catch (e) {
    throw new Error('Erro ao abrir a planilha: ' + e.message);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _norm_(s) { // Usado apenas para mapear cabeçalhos
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\./g, ' ').replace(/\//g, ' ').replace(/-/g, ' ')
    .trim().toLowerCase().replace(/\s+/g, '_')
    .replace(/[^a-z0-9_]/g, '_').replace(/_+/g, '_');
}

function _toNumber_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').trim();
  if (!s) return 0;
  const clean = s.replace(/R\$/g, '').replace(/\./g, '').replace(/,/g, '.').replace(/[^\d.-]/g, '');
  const n = parseFloat(clean);
  return isNaN(n) ? 0 : n;
}

function _toInt_(v) {
  const s = String(v ?? '').replace(/[^\d-]/g, '').trim();
  const n = parseInt(s, 10);
  return isNaN(n) ? null : n;
}

function _asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v || '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/); // yyyy-mm-dd
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);   // dd/mm/aaaa
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function _fmtBR_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return Utilities.formatDate(d, TZ, 'dd/MM/yyyy');
}

function _fmtBRDateTime_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return Utilities.formatDate(d, TZ, 'dd/MM/yyyy HH:mm:ss');
}

function _fmtSortableDate_(d) {
  if (!(d instanceof Date) || isNaN(d)) return '';
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}


// ====== LEITURA (MODIFICADA) ======
function _readTable_() {
  const ss = _openSS_();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Aba não encontrada: ' + SHEET_NAME);
  
  const timestampValue = sh.getRange('H2').getDisplayValue();

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  
  if (lastRow < HEADER_ROW) return { headers: [], rows: [], timestampValue: timestampValue };

  const MAX_DATA_ROWS = 5000;
  const totalRowsAvailable = lastRow - HEADER_ROW + 1; 
  const numRowsToRead = Math.min(totalRowsAvailable, MAX_DATA_ROWS + 1);

  if (numRowsToRead <= 1) return { headers: [], rows: [], timestampValue: timestampValue };

  const range = sh.getRange(HEADER_ROW, 1, numRowsToRead, lastCol);
  const values = range.getValues();
  if (!values || values.length < 2) return { headers: [], rows: [], timestampValue: timestampValue }; 

  const headers = values[0];
  const rows = values.slice(1).filter(r => r.some(c => c !== '' && c !== null));
  
  return { headers, rows, timestampValue };
}

// ====== MAPEAMENTO DE CABEÇALHOS (NENHUMA MUDANÇA NECESSÁRIA) ======
function _buildHeaderIndex_(headers) {
  const norms = headers.map(_norm_);
  function find(...aliases) {
    for (let i = 0; i < norms.length; i++) if (aliases.includes(norms[i])) return i;
    for (let i = 0; i < norms.length; i++) {
      const n = norms[i];
      if (aliases.some(a => n.includes(a))) return i;
    }
    return -1;
  }
  const idx = {
    cartela:      find('cartela'),
    cliente:      find('cliente'),
    descricao:    find('descricao', 'descri', 'descricao_item', 'produto', 'item'),
    tamanho:      find('tamanho', 'tam'),
    ord_compra:   find('ord_compra', 'ord__compra', 'ordem_compra', 'ordem_de_compra', 'oc'),
    qtd_aberta:   find('qtd_aberta', 'qtde_aberta', 'quantidade_aberta', 'saldo', 'saldo_aberto'),
    data_receb:   find('data_receb', 'dt_receb', 'recebimento'),
    dt_entrega:   find('dt_entrega', 'data_entrega', 'entrega', 'previsao_entrega'),
    data_sistema: find('data_sistema'),
    prazo:        find('prazo')
  };
  const missing = Object.entries(idx).filter(([, v]) => v < 0).map(([k]) => k);
  return { idx, missing, norms, headers };
}

// ====== LINHA -> ITEM (sem alteração) ======
function _rowsToItems_(rows, idx) {
  const out = [];
  for (const r of rows) {
    const prazoStr = r[idx.prazo];
    const prazoNum = _toInt_(prazoStr);

    const dtEntregaDate = _asDate_((r[idx.dt_entrega]));

    const it = {
      cartela:         String(r[idx.cartela]      ?? '').trim(),
      cliente:         String(r[idx.cliente]      ?? '').trim(),
      descricao:       String(r[idx.descricao]     ?? '').trim(),
      tamanho:         String(r[idx.tamanho]      ?? '').trim(),
      ord_compra:      String(r[idx.ord_compra]   ?? '').trim(),
      qtd_aberta:      _toNumber_(r[idx.qtd_aberta]),
      data_receb_br:   _fmtBR_(_asDate_(r[idx.data_receb])),
      dt_entrega_br:   _fmtBR_(dtEntregaDate),
      data_sistema_br: _fmtBR_(_asDate_(r[idx.data_sistema])),
      prazo:           String(prazoStr ?? '').trim(),
      prazo_num:       prazoNum,
      dt_entrega_sortable: _fmtSortableDate_(dtEntregaDate)
    };
    out.push(it);
  }
  return out;
}


// ====== API: FETCH ALL DATA (com contador) ======
function fetchAllData(cacheBuster) {
  try {
    if (cacheBuster) {
      console.log('fetchAllData: Cache buster recebido: ' + cacheBuster);
    }
    
    console.log('fetchAllData: Iniciando leitura da tabela...');
    
    const { headers, rows, timestampValue } = _readTable_();
    
    const { idx, missing, norms } = _buildHeaderIndex_(headers);
    
    if (missing.length) {
      return { 
        ok: false, 
        error: 'Não encontrei as colunas: ' + missing.join(', '),
        info: { headers_detectados: headers, headers_norm: norms } 
      };
    }

    console.log('fetchAllData: Convertendo ' + rows.length + ' linhas para itens...');
    const all = _rowsToItems_(rows, idx);
    console.log('fetchAllData: Enviando ' + all.length + ' itens para o cliente.');

    let displayTimestamp = '';
    if (timestampValue) {
      displayTimestamp = 'Atualizado: ' + String(timestampValue).trim();
    } else {
      displayTimestamp = 'Data de atualização não informada (H2).';
    }

    const now = new Date();
    const requestTimestamp = _fmtBRDateTime_(now);
    
    const accessCount = _getAccessCount_();

    return {
      ok: true,
      items: all,
      meta: {
        updated_at_display: displayTimestamp,
        request_timestamp: requestTimestamp,
        access_count_today: accessCount,
        cache_buster: cacheBuster || 'none',
        version: APP_VERSION,
        author: 'Johnny',
        rows_read: all.length
      }
    };
  } catch (err) {
    console.error('Erro em fetchAllData: ' + err.message + ' Stack: ' + err.stack);
    return { 
      ok: false, 
      error: 'fetchAllData: ' + err.message 
    };
  }
}

// ====== CONTADOR DE ACESSOS (sem alteração) ======
function _getAccessCount_() {
  let lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    
    const lastResetDate = props.getProperty('LAST_RESET_DATE');
    let accessCount = 0;
    
    if (lastResetDate != today) {
      accessCount = 1;
      props.setProperties({
        'LAST_RESET_DATE': today,
        'ACCESS_COUNT_TODAY': accessCount
      });
      console.log('Novo dia, contador de acessos resetado para 1.');
    } else {
      accessCount = (Number(props.getProperty('ACCESS_COUNT_TODAY')) || 0) + 1;
      props.setProperty('ACCESS_COUNT_TODAY', accessCount);
    }
    
    return accessCount;

  } catch (e) {
    console.error('Falha ao obter lock ou incrementar contador: ' + e.message);
    return PropertiesService.getScriptProperties().getProperty('ACCESS_COUNT_TODAY') || 0;
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}
