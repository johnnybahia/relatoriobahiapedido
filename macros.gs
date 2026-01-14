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
      .addItem('🖨️ Imprimir Relatório por OC(s)', 'mostrarDialogoOCs')
      .addSeparator()
      .addItem('🔧 Testar Macro (Debug)', 'testarMacroDebug')
      .addSeparator()
      .addSubMenu(ui.createMenu('💰 Faturamento')
          .addItem('🔄 Atualizar Faturamento Agora', 'atualizarFaturamentoManual')
          .addItem('⚙️ Configurar Triggers Automáticos (8h e 19h)', 'criarTriggersAutomaticos')
          .addItem('🧹 Limpar Faturamento do Dia', 'limparFaturamentoDia'))
      .addToUi();
}

/**
 * Função de teste para debug - teste rápido com uma OC
 */
function testarMacroDebug() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Teste de Debug',
    'Digite uma OC para teste:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.CANCEL) {
    return;
  }

  const ocTeste = result.getResponseText().trim();
  if (!ocTeste) {
    ui.alert('Por favor, digite uma OC válida.');
    return;
  }

  console.log('=== INÍCIO DO TESTE DE DEBUG ===');
  console.log('OC para teste:', ocTeste);

  try {
    const resultado = processarMultiplasOCs([ocTeste]);
    console.log('=== RESULTADO DO TESTE ===');
    console.log('Success:', resultado.success);
    console.log('Message:', resultado.message);

    if (resultado.success) {
      ui.alert('✅ Teste concluído com sucesso!\n\n' + resultado.message + '\n\nVerifique o log de execuções (Ctrl+Enter) para mais detalhes.');
    } else {
      ui.alert('❌ Teste falhou!\n\n' + resultado.message + '\n\nVerifique o log de execuções (Ctrl+Enter) para mais detalhes.');
    }
  } catch (error) {
    console.error('=== ERRO NO TESTE ===');
    console.error(error.toString());
    console.error(error.stack);
    ui.alert('❌ Erro no teste:\n\n' + error.message + '\n\nVerifique o log de execuções (Ctrl+Enter) para detalhes completos.');
  }

  console.log('=== FIM DO TESTE DE DEBUG ===');
}

/**
 * Mostra diálogo HTML para inserir múltiplas OCs de forma prática
 */
function mostrarDialogoOCs() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            background: #f5f5f5;
          }
          .container {
            background: white;
            padding: 25px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            max-width: 500px;
            margin: 0 auto;
          }
          h2 {
            color: #333;
            margin-top: 0;
            border-bottom: 3px solid #4CAF50;
            padding-bottom: 10px;
          }
          .info {
            background: #e3f2fd;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
            border-left: 4px solid #2196F3;
          }
          .info strong { color: #1976D2; }
          label {
            display: block;
            margin-bottom: 8px;
            color: #555;
            font-weight: bold;
          }
          textarea {
            width: 100%;
            min-height: 150px;
            padding: 12px;
            border: 2px solid #ddd;
            border-radius: 5px;
            font-size: 14px;
            font-family: monospace;
            box-sizing: border-box;
            resize: vertical;
          }
          textarea:focus {
            outline: none;
            border-color: #4CAF50;
          }
          .button-group {
            display: flex;
            gap: 10px;
            margin-top: 20px;
          }
          button {
            flex: 1;
            padding: 12px 24px;
            font-size: 16px;
            border: none;
            border-radius: 5px;
            cursor: pointer;
            font-weight: bold;
            transition: all 0.3s;
          }
          #btnGerar {
            background: #4CAF50;
            color: white;
          }
          #btnGerar:hover {
            background: #45a049;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(76, 175, 80, 0.3);
          }
          #btnCancelar {
            background: #f44336;
            color: white;
          }
          #btnCancelar:hover {
            background: #da190b;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(244, 67, 54, 0.3);
          }
          button:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
          }
          .exemplo {
            font-size: 12px;
            color: #666;
            margin-top: 8px;
            font-style: italic;
          }
          #status {
            margin-top: 15px;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
            font-weight: bold;
            display: none;
          }
          .success { background: #d4edda; color: #155724; display: block; }
          .error { background: #f8d7da; color: #721c24; display: block; }
          .loading { background: #fff3cd; color: #856404; display: block; }
        </style>
      </head>
      <body>
        <div class="container">
          <h2>📋 Gerador de Relatórios por OC</h2>

          <div class="info">
            <strong>💡 Dica:</strong> Insira uma ou várias Ordens de Compra (OC).<br>
            Use uma OC por linha ou separe por vírgula.
          </div>

          <label for="inputOCs">Digite as OCs:</label>
          <textarea id="inputOCs" placeholder="Exemplo:&#10;12345&#10;67890&#10;11223&#10;&#10;Ou: 12345, 67890, 11223"></textarea>
          <div class="exemplo">Exemplo: 12345, 67890 ou uma por linha</div>

          <div id="status"></div>

          <div class="button-group">
            <button id="btnCancelar" onclick="google.script.host.close()">❌ Cancelar</button>
            <button id="btnGerar" onclick="gerarRelatorios()">✅ Gerar Relatório(s)</button>
          </div>
        </div>

        <script>
          function gerarRelatorios() {
            const input = document.getElementById('inputOCs').value.trim();
            const status = document.getElementById('status');
            const btnGerar = document.getElementById('btnGerar');

            if (!input) {
              status.className = 'error';
              status.textContent = '⚠️ Por favor, insira pelo menos uma OC!';
              return;
            }

            // Separa OCs por vírgula, quebra de linha ou ponto e vírgula
            const ocs = input
              .split(/[,;\\n]+/)
              .map(oc => oc.trim())
              .filter(oc => oc !== '');

            if (ocs.length === 0) {
              status.className = 'error';
              status.textContent = '⚠️ Nenhuma OC válida encontrada!';
              return;
            }

            status.className = 'loading';
            status.textContent = '⏳ Gerando relatório(s)... Aguarde.';
            btnGerar.disabled = true;

            // Chama a função do servidor
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  status.className = 'success';
                  status.textContent = '✅ ' + result.message;
                  setTimeout(function() {
                    google.script.host.close();
                  }, 1500);
                } else {
                  status.className = 'error';
                  status.textContent = '❌ ' + result.message;
                  btnGerar.disabled = false;
                }
              })
              .withFailureHandler(function(error) {
                status.className = 'error';
                status.textContent = '❌ Erro: ' + error.message;
                btnGerar.disabled = false;
              })
              .processarMultiplasOCs(ocs);
          }

          // Atalho Enter com Ctrl para gerar
          document.getElementById('inputOCs').addEventListener('keydown', function(e) {
            if (e.ctrlKey && e.key === 'Enter') {
              gerarRelatorios();
            }
          });
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(550)
    .setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gerador de Relatórios');
}

/**
 * Processa múltiplas OCs e gera o relatório
 */
function processarMultiplasOCs(ocs) {
  try {
    console.log('Iniciando processarMultiplasOCs com', ocs.length, 'OCs');

    if (!ocs || ocs.length === 0) {
      console.log('Erro: Nenhuma OC fornecida');
      return { success: false, message: 'Nenhuma OC fornecida' };
    }

    console.log('OCs recebidas:', ocs.join(', '));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    console.log('Planilha ativa:', sheet.getName());

    // Otimização: lê dados uma única vez
    console.log('Lendo dados da planilha...');
    const dadosCache = lerDadosOtimizado(sheet, ss);

    if (!dadosCache.success) {
      console.log('Erro ao ler dados:', dadosCache.message);
      return { success: false, message: dadosCache.message };
    }

    console.log('Dados lidos com sucesso. Total de linhas:', dadosCache.dados.length);

    // Gera relatório para as OCs
    console.log('Gerando HTML do relatório...');
    const resultado = gerarRelatorioMultiplasOCs(ocs, dadosCache);

    if (!resultado || !resultado.html) {
      console.log('Erro: Nenhum item encontrado para as OCs fornecidas');
      return { success: false, message: 'Nenhum item encontrado para as OCs fornecidas' };
    }

    console.log('OCs encontradas:', resultado.ocsEncontradas.length);
    console.log('OCs não encontradas:', resultado.ocsNaoEncontradas.length);
    console.log('Tamanho do HTML gerado:', resultado.html.length, 'caracteres');

    // Verifica se o HTML não é muito grande (limite do Google Apps Script)
    if (resultado.html.length > 1000000) {
      console.log('AVISO: HTML muito grande, pode causar problemas');
      return {
        success: false,
        message: 'Relatório muito grande. Tente gerar menos OCs por vez (máximo 5-10).'
      };
    }

    // Mostra o relatório
    console.log('Criando janela do relatório...');
    try {
      const output = HtmlService.createHtmlOutput(resultado.html)
        .setWidth(1200)
        .setHeight(700);

      SpreadsheetApp.getUi().showModalDialog(output, 'Relatório de Produção');
      console.log('Janela do relatório criada com sucesso');
    } catch (e) {
      console.error('Erro ao criar janela do relatório:', e.toString());
      return {
        success: false,
        message: 'Erro ao exibir relatório: ' + e.message + '. Tente gerar menos OCs por vez.'
      };
    }

    // Monta mensagem com informações sobre OCs encontradas e não encontradas
    let msg = '';
    if (resultado.ocsEncontradas.length === 1 && resultado.ocsNaoEncontradas.length === 0) {
      msg = 'Relatório gerado com sucesso!';
    } else if (resultado.ocsNaoEncontradas.length === 0) {
      msg = `Relatórios gerados para ${resultado.ocsEncontradas.length} OC(s)!`;
    } else {
      msg = `Relatórios gerados para ${resultado.ocsEncontradas.length} OC(s). ⚠️ ${resultado.ocsNaoEncontradas.length} OC(s) não encontrada(s).`;
    }

    console.log('Processo concluído com sucesso:', msg);
    return { success: true, message: msg };

  } catch (error) {
    console.error('ERRO em processarMultiplasOCs:', error.toString());
    console.error('Stack trace:', error.stack);
    return {
      success: false,
      message: 'Erro ao processar: ' + error.message + '. Verifique o log de execuções (Ctrl+Enter).'
    };
  }
}

/**
 * Lê dados de forma otimizada (uma única vez)
 */
function lerDadosOtimizado(sheet, ss) {
  try {
    console.log('lerDadosOtimizado: Iniciando...');

    // Otimização: usa getRange específico ao invés de getDataRange
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    console.log(`lerDadosOtimizado: lastRow=${lastRow}, lastCol=${lastCol}`);

    if (lastRow < 3) {
      console.log('lerDadosOtimizado: Erro - Planilha sem dados suficientes');
      return { success: false, message: 'Planilha sem dados suficientes' };
    }

    // Lê apenas até a última linha com dados (não toda a planilha)
    console.log('lerDadosOtimizado: Lendo range da planilha...');
    const dados = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    console.log(`lerDadosOtimizado: ${dados.length} linhas lidas`);

    const INDICE_CABECALHO = 2;
    const cabecalho = dados[INDICE_CABECALHO].map(c => String(c).trim().toUpperCase());
    console.log('lerDadosOtimizado: Cabeçalhos:', cabecalho.slice(0, 5).join(', '), '...');

    // Mapear índices
    const mapa = {};
    for (let key in COLUNAS) {
      let index = cabecalho.indexOf(COLUNAS[key].toUpperCase());
      if (index === -1) {
        index = cabecalho.findIndex(c => c.includes(COLUNAS[key].toUpperCase()));
      }
      mapa[key] = index;
    }
    console.log('lerDadosOtimizado: Mapa de colunas criado');

    // Verificação de QTD ABERTA
    if (mapa.QTD_ABERTA === -1) {
      let indexSemPonto = cabecalho.indexOf("QTD ABERTA");
      if (indexSemPonto === -1) {
        indexSemPonto = cabecalho.findIndex(c => c.includes("QTD") && c.includes("ABERTA"));
      }
      if (indexSemPonto > -1) {
        mapa.QTD_ABERTA = indexSemPonto;
      }
    }

    // Buscar marcas na aba MARCAS
    console.log('lerDadosOtimizado: Buscando aba MARCAS...');
    const sheetMarcas = ss.getSheetByName(ABA_MARCAS_NOME);
    let marcasMap = {};

    if (sheetMarcas) {
      console.log('lerDadosOtimizado: Aba MARCAS encontrada, lendo dados...');
      const dadosMarcas = sheetMarcas.getDataRange().getValues();
      const headerMarcas = dadosMarcas[0].map(c => String(c).toUpperCase().trim());

      let colIndexOcMarca = headerMarcas.indexOf("ORDEM DE COMPRA");
      if (colIndexOcMarca === -1) colIndexOcMarca = headerMarcas.indexOf("ORD. COMPRA");
      if (colIndexOcMarca === -1) colIndexOcMarca = 0;

      let colIndexNomeMarca = headerMarcas.indexOf("MARCA");
      if (colIndexNomeMarca === -1) colIndexNomeMarca = 1;

      for (let i = 1; i < dadosMarcas.length; i++) {
        const oc = String(dadosMarcas[i][colIndexOcMarca]).trim();
        const marca = dadosMarcas[i][colIndexNomeMarca];
        if (oc) marcasMap[oc] = marca;
      }
      console.log(`lerDadosOtimizado: ${Object.keys(marcasMap).length} marcas carregadas`);
    } else {
      console.log('lerDadosOtimizado: Aba MARCAS não encontrada');
    }

    console.log('lerDadosOtimizado: Concluído com sucesso');
    return {
      success: true,
      dados: dados,
      mapa: mapa,
      marcasMap: marcasMap,
      INDICE_CABECALHO: INDICE_CABECALHO
    };

  } catch (error) {
    console.error('lerDadosOtimizado: ERRO -', error.toString());
    console.error('lerDadosOtimizado: Stack -', error.stack);
    return { success: false, message: 'Erro ao ler dados: ' + error.toString() };
  }
}

/**
 * Gera HTML do relatório para múltiplas OCs
 */
function gerarRelatorioMultiplasOCs(ocs, dadosCache) {
  try {
    console.log('gerarRelatorioMultiplasOCs: Iniciando...');
    console.log('gerarRelatorioMultiplasOCs: OCs a processar:', ocs.join(', '));

    const { dados, mapa, marcasMap, INDICE_CABECALHO } = dadosCache;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataHoje = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm");

    // Agrupa itens por OC
    console.log('gerarRelatorioMultiplasOCs: Agrupando itens por OC...');
    const itensPorOC = {};
    const clientesPorOC = {};

    let linhasProcessadas = 0;
    for (let i = INDICE_CABECALHO + 1; i < dados.length; i++) {
      const linha = dados[i];
      if (linha.length <= mapa.ORDEM_COMPRA) continue;

      const ocLinha = String(linha[mapa.ORDEM_COMPRA]).trim();

      if (ocs.includes(ocLinha)) {
        if (!itensPorOC[ocLinha]) {
          itensPorOC[ocLinha] = [];
          clientesPorOC[ocLinha] = mapa.CLIENTE > -1 ? linha[mapa.CLIENTE] : "";
          console.log(`gerarRelatorioMultiplasOCs: OC ${ocLinha} encontrada`);
        }

        itensPorOC[ocLinha].push({
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
        linhasProcessadas++;
      }
    }

    console.log(`gerarRelatorioMultiplasOCs: ${linhasProcessadas} itens encontrados`);

    // Verifica quais OCs foram encontradas e quais não
    const ocsEncontradas = Object.keys(itensPorOC);
    const ocsNaoEncontradas = ocs.filter(oc => !ocsEncontradas.includes(oc));

    console.log('gerarRelatorioMultiplasOCs: OCs encontradas:', ocsEncontradas.join(', '));
    console.log('gerarRelatorioMultiplasOCs: OCs não encontradas:', ocsNaoEncontradas.join(', '));

    if (ocsEncontradas.length === 0) {
      console.log('gerarRelatorioMultiplasOCs: Nenhuma OC encontrada!');
      return { html: null, ocsEncontradas: [], ocsNaoEncontradas: ocsNaoEncontradas };
    }

    console.log('gerarRelatorioMultiplasOCs: Gerando HTML...');
  } catch (error) {
    console.error('gerarRelatorioMultiplasOCs: ERRO ao processar dados -', error.toString());
    throw error;
  }

  // Gera HTML com LOGO e destaque da OC - OTIMIZADO PARA LASER P&B VERTICAL
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
          font-family: Arial, sans-serif;
          font-size: 10px;
          padding: 15px;
          background: #fff;
          color: #000;
        }
        .oc-section {
          margin-bottom: 30px;
          page-break-after: always;
          border: 3px double #000;
          padding: 15px;
        }
        .header {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 15px;
          margin-bottom: 15px;
          padding-bottom: 10px;
          border-bottom: 1px solid #000;
        }
        .header img {
          max-width: 80px;
          height: auto;
        }
        .header h2 {
          margin: 0;
          color: #000;
          font-size: 14px;
          text-align: center;
        }
        .oc-destaque {
          border: 4px double #000;
          padding: 12px;
          margin: 12px 0;
          text-align: center;
          font-size: 20px;
          font-weight: bold;
          letter-spacing: 1px;
          background: #fff;
          color: #000;
        }
        .info-line {
          margin-bottom: 6px;
          padding: 5px 0;
          border-bottom: 1px dotted #666;
        }
        .info-line strong {
          font-weight: bold;
          text-transform: uppercase;
        }
        .destaque-marca {
          font-size: 12px;
          margin-top: 8px;
          font-weight: bold;
          padding: 8px;
          border: 2px solid #000;
          text-align: center;
          background: #fff;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin-top: 12px;
        }
        th, td {
          border: 1px solid #000;
          padding: 5px 3px;
          text-align: center;
          font-size: 9px;
          background: #fff;
        }
        th {
          font-weight: bold;
          text-transform: uppercase;
          border: 2px solid #000;
        }
        .text-left {
          text-align: left;
          font-size: 8px;
        }
        .col-pedido { width: 8%; }
        .col-cod-cliente { width: 8%; }
        .col-cod-marfim { width: 8%; }
        .col-descricao { width: 25%; }
        .col-tamanho { width: 6%; }
        .col-qtd { width: 7%; }
        .col-lotes { width: 8%; }
        .col-os { width: 8%; }
        .col-data { width: 8%; }
        .col-prazo { width: 6%; }
        .btn-print {
          padding: 10px 25px;
          background: #333;
          color: white;
          border: 2px solid #000;
          cursor: pointer;
          font-size: 14px;
          font-weight: bold;
          margin-bottom: 15px;
          display: block;
          transition: all 0.3s;
        }
        .btn-print:hover {
          background: #555;
        }
        .rodape {
          margin-top: 20px;
          padding-top: 10px;
          border-top: 1px solid #000;
          text-align: center;
          font-size: 9px;
        }
        .aviso-ocs {
          background: #fff3cd;
          border: 2px solid #856404;
          padding: 15px;
          margin-bottom: 20px;
          border-radius: 5px;
        }
        .aviso-ocs h3 {
          color: #856404;
          margin-bottom: 10px;
          font-size: 14px;
        }
        .aviso-ocs ul {
          margin-left: 20px;
          color: #856404;
          font-size: 12px;
        }
        .aviso-ocs li {
          margin-bottom: 5px;
        }
        @media print {
          .btn-print { display: none !important; }
          .aviso-ocs { display: none !important; }
          body {
            padding: 0.5cm;
            font-size: 9px;
          }
          .header img {
            max-width: 60px;
          }
          .oc-section {
            page-break-after: always;
            padding: 10px;
          }
          .oc-section:last-child {
            page-break-after: auto;
          }
          .oc-destaque {
            font-size: 18px;
            padding: 10px;
          }
          @page {
            size: portrait;
            margin: 1cm;
          }
        }
        @media screen {
          body { max-width: 21cm; margin: 0 auto; }
        }
      </style>
    </head>
    <body>
      <button class="btn-print" onclick="window.print()">🖨️ IMPRIMIR RELATÓRIO(S)</button>
  `;

  // Mostra aviso se houver OCs não encontradas
  if (ocsNaoEncontradas.length > 0) {
    html += `
      <div class="aviso-ocs">
        <h3>⚠️ Atenção: ${ocsNaoEncontradas.length} OC(s) não encontrada(s)</h3>
        <p><strong>As seguintes OCs não foram encontradas na planilha:</strong></p>
        <ul>
    `;
    ocsNaoEncontradas.forEach(oc => {
      html += `<li>OC: <strong>${oc}</strong></li>`;
    });
    html += `
        </ul>
        <p style="margin-top: 10px;"><em>Os relatórios das OCs encontradas (${ocsEncontradas.length}) serão gerados normalmente.</em></p>
      </div>
    `;
  }

  // Gera uma seção para cada OC
  ocsEncontradas.forEach((oc, index) => {
    const itens = itensPorOC[oc];
    const cliente = clientesPorOC[oc];
    const marca = marcasMap[oc] || "N/A";

    html += `
      <div class="oc-section">
        <div class="header">
          <img src="https://i.ibb.co/FGGjdsM/LOGO-MARFIM.jpg" alt="Logo MARFIM" onerror="this.style.display='none'">
          <h2>📋 RELATÓRIO DE PRODUÇÃO</h2>
        </div>

        <div class="oc-destaque">
          ORDEM DE COMPRA: ${oc}
        </div>

        <div class="info-line">
          <strong>Cliente:</strong> ${cliente}
        </div>

        <div class="info-line">
          <strong>Data:</strong> ${dataHoje}
        </div>

        <div class="destaque-marca">
          MARCA: ${String(marca).toUpperCase()}
        </div>

        <table>
          <thead>
            <tr>
              <th class="col-pedido">PEDIDO</th>
              <th class="col-cod-cliente">CÓD.<br>CLIENTE</th>
              <th class="col-cod-marfim">CÓD.<br>MARFIM</th>
              <th class="col-descricao">DESCRIÇÃO</th>
              <th class="col-tamanho">TAM.</th>
              <th class="col-qtd">QTD.<br>ABERTA</th>
              <th class="col-lotes">LOTES</th>
              <th class="col-os">CÓD.<br>OS</th>
              <th class="col-data">DT.<br>REC.</th>
              <th class="col-data">DT.<br>ENT.</th>
              <th class="col-prazo">PRAZO</th>
            </tr>
          </thead>
          <tbody>
    `;

    itens.forEach(item => {
      html += `
        <tr>
          <td class="col-pedido">${item.pedido}</td>
          <td class="col-cod-cliente">${item.codCliente}</td>
          <td class="col-cod-marfim">${item.codMarfim}</td>
          <td class="col-descricao text-left">${item.descricao}</td>
          <td class="col-tamanho">${item.tamanho}</td>
          <td class="col-qtd">${item.qtdAberta}</td>
          <td class="col-lotes">${item.lotes}</td>
          <td class="col-os">${item.codOs}</td>
          <td class="col-data">${item.dtRec}</td>
          <td class="col-data">${item.dtEnt}</td>
          <td class="col-prazo">${item.prazo}</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>

        <div class="rodape">
          <p><strong>Total de itens:</strong> ${itens.length} | <strong>Gerado em:</strong> ${dataHoje}</p>
        </div>
      </div>
    `;
  });

  html += `
    </body>
    </html>
  `;

  console.log('gerarRelatorioMultiplasOCs: HTML gerado com sucesso!');
  console.log(`gerarRelatorioMultiplasOCs: Tamanho total: ${html.length} caracteres`);

  return {
    html: html,
    ocsEncontradas: ocsEncontradas,
    ocsNaoEncontradas: ocsNaoEncontradas
  };
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

// ====== SISTEMA DE FATURAMENTO ======

/**
 * Lê dados da aba "Dados1" (Ordem de Compra, Valor, Cliente)
 */
function lerDados1() {
  try {
    const ss = _openSS_();
    const sheet = ss.getSheetByName('Dados1');

    if (!sheet) {
      console.error('Aba Dados1 não encontrada');
      return { success: false, error: 'Aba Dados1 não encontrada' };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { success: true, dados: [] };
    }

    const range = sheet.getRange(2, 1, lastRow - 1, 3);
    const values = range.getValues();

    const dados = [];
    for (let i = 0; i < values.length; i++) {
      const oc = String(values[i][0]).trim();
      const valor = _toNumber_(values[i][1]);
      const cliente = String(values[i][2]).trim();

      if (oc && cliente) {
        dados.push({ oc, valor, cliente });
      }
    }

    console.log(`lerDados1: ${dados.length} registros lidos`);
    return { success: true, dados };

  } catch (error) {
    console.error('Erro em lerDados1:', error);
    return { success: false, error: error.toString() };
  }
}

/**
 * Salva snapshot dos dados atuais para comparação futura
 */
function salvarSnapshot() {
  try {
    console.log('Salvando snapshot dos dados...');
    const resultado = lerDados1();

    if (!resultado.success) {
      console.error('Erro ao ler dados para snapshot:', resultado.error);
      return;
    }

    const props = PropertiesService.getScriptProperties();
    const now = new Date();
    const timestamp = Utilities.formatDate(now, TZ, 'dd/MM/yyyy HH:mm:ss');

    props.setProperty('SNAPSHOT_DADOS', JSON.stringify(resultado.dados));
    props.setProperty('SNAPSHOT_TIMESTAMP', timestamp);

    console.log(`Snapshot salvo: ${resultado.dados.length} registros em ${timestamp}`);

  } catch (error) {
    console.error('Erro ao salvar snapshot:', error);
  }
}

/**
 * Detecta faturamento comparando dados atuais com snapshot anterior
 */
function detectarFaturamento() {
  try {
    console.log('Detectando faturamento...');
    const props = PropertiesService.getScriptProperties();

    // Lê snapshot anterior
    const snapshotJson = props.getProperty('SNAPSHOT_DADOS');
    if (!snapshotJson) {
      console.log('Nenhum snapshot anterior encontrado. Salvando primeiro snapshot...');
      salvarSnapshot();
      return { faturados: [], timestamp: null };
    }

    const dadosAntigos = JSON.parse(snapshotJson);
    const timestampAnterior = props.getProperty('SNAPSHOT_TIMESTAMP');

    // Lê dados atuais
    const resultado = lerDados1();
    if (!resultado.success) {
      console.error('Erro ao ler dados atuais');
      return { faturados: [], timestamp: timestampAnterior };
    }

    const dadosAtuais = resultado.dados;

    // Cria mapa dos dados atuais por OC
    const mapaAtual = {};
    dadosAtuais.forEach(d => {
      mapaAtual[d.oc] = d;
    });

    // Detecta faturamento
    const faturados = [];

    dadosAntigos.forEach(dadoAntigo => {
      const dadoAtual = mapaAtual[dadoAntigo.oc];

      if (!dadoAtual) {
        // OC sumiu = foi totalmente faturada
        faturados.push({
          oc: dadoAntigo.oc,
          cliente: dadoAntigo.cliente,
          valorFaturado: dadoAntigo.valor
        });
        console.log(`OC ${dadoAntigo.oc} foi faturada (sumiu): R$ ${dadoAntigo.valor}`);

      } else if (dadoAtual.valor < dadoAntigo.valor) {
        // Valor diminuiu = faturamento parcial
        const valorFaturado = dadoAntigo.valor - dadoAtual.valor;
        faturados.push({
          oc: dadoAntigo.oc,
          cliente: dadoAntigo.cliente,
          valorFaturado: valorFaturado
        });
        console.log(`OC ${dadoAntigo.oc} faturamento parcial: R$ ${valorFaturado}`);
      }
    });

    console.log(`Total de faturamentos detectados: ${faturados.length}`);

    // Salva faturamento detectado
    if (faturados.length > 0) {
      const faturamentoAtual = props.getProperty('FATURAMENTO_DIA');
      const faturamentoDia = faturamentoAtual ? JSON.parse(faturamentoAtual) : [];

      faturados.forEach(f => faturamentoDia.push(f));

      props.setProperty('FATURAMENTO_DIA', JSON.stringify(faturamentoDia));
    }

    // Salva novo snapshot
    salvarSnapshot();

    return { faturados, timestamp: timestampAnterior };

  } catch (error) {
    console.error('Erro ao detectar faturamento:', error);
    return { faturados: [], timestamp: null };
  }
}

/**
 * Limpa faturamento do dia (executar no início do dia)
 */
function limparFaturamentoDia() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('FATURAMENTO_DIA', JSON.stringify([]));
    console.log('Faturamento do dia limpo');
  } catch (error) {
    console.error('Erro ao limpar faturamento:', error);
  }
}

/**
 * Função manual para atualizar faturamento
 */
function atualizarFaturamentoManual() {
  console.log('=== ATUALIZAÇÃO MANUAL DE FATURAMENTO ===');
  const resultado = detectarFaturamento();
  console.log(`Faturamentos detectados: ${resultado.faturados.length}`);
  console.log('=== FIM DA ATUALIZAÇÃO ===');

  const ui = SpreadsheetApp.getUi();
  if (resultado.faturados.length > 0) {
    ui.alert(`✅ Faturamento atualizado!\n\n${resultado.faturados.length} movimentação(ões) detectada(s).`);
  } else {
    ui.alert('ℹ️ Nenhuma mudança detectada desde a última verificação.');
  }
}

/**
 * Processa dados para o card "Pedidos a Faturar"
 */
function getPedidosAFaturar() {
  try {
    const resultado = lerDados1();
    if (!resultado.success || resultado.dados.length === 0) {
      return {
        sucesso: true,
        dados: [],
        timestamp: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss')
      };
    }

    const ss = _openSS_();
    const sheetMarcas = ss.getSheetByName(ABA_MARCAS_NOME);
    const marcasMap = {};

    if (sheetMarcas) {
      const dadosMarcas = sheetMarcas.getDataRange().getValues();
      for (let i = 1; i < dadosMarcas.length; i++) {
        const oc = String(dadosMarcas[i][0]).trim();
        const marca = dadosMarcas[i][1];
        if (oc) marcasMap[oc] = marca;
      }
    }

    // Agrupa por cliente
    const porCliente = {};

    resultado.dados.forEach(item => {
      const marca = marcasMap[item.oc] || 'N/A';
      const chave = `${item.cliente}|${marca}`;

      if (!porCliente[chave]) {
        porCliente[chave] = {
          cliente: item.cliente,
          marca: marca,
          valor: 0
        };
      }

      porCliente[chave].valor += item.valor;
    });

    const dados = Object.values(porCliente).sort((a, b) => b.valor - a.valor);

    return {
      sucesso: true,
      dados: dados,
      timestamp: Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss')
    };

  } catch (error) {
    console.error('Erro em getPedidosAFaturar:', error);
    return { sucesso: false, erro: error.toString() };
  }
}

/**
 * Processa dados para o card "Faturamento do Dia"
 */
function getFaturamentoDia() {
  try {
    const props = PropertiesService.getScriptProperties();
    const faturamentoJson = props.getProperty('FATURAMENTO_DIA');

    if (!faturamentoJson) {
      return {
        sucesso: true,
        dados: [],
        timestamp: null
      };
    }

    const faturados = JSON.parse(faturamentoJson);

    if (faturados.length === 0) {
      return {
        sucesso: true,
        dados: [],
        timestamp: props.getProperty('SNAPSHOT_TIMESTAMP')
      };
    }

    const ss = _openSS_();
    const sheetMarcas = ss.getSheetByName(ABA_MARCAS_NOME);
    const marcasMap = {};

    if (sheetMarcas) {
      const dadosMarcas = sheetMarcas.getDataRange().getValues();
      for (let i = 1; i < dadosMarcas.length; i++) {
        const oc = String(dadosMarcas[i][0]).trim();
        const marca = dadosMarcas[i][1];
        if (oc) marcasMap[oc] = marca;
      }
    }

    // Agrupa por cliente e marca
    const porCliente = {};

    faturados.forEach(item => {
      const marca = marcasMap[item.oc] || 'N/A';
      const chave = `${item.cliente}|${marca}`;

      if (!porCliente[chave]) {
        porCliente[chave] = {
          cliente: item.cliente,
          marca: marca,
          valor: 0
        };
      }

      porCliente[chave].valor += item.valorFaturado;
    });

    const dados = Object.values(porCliente).sort((a, b) => b.valor - a.valor);

    return {
      sucesso: true,
      dados: dados,
      timestamp: props.getProperty('SNAPSHOT_TIMESTAMP')
    };

  } catch (error) {
    console.error('Erro em getFaturamentoDia:', error);
    return { sucesso: false, erro: error.toString() };
  }
}

/**
 * Cria triggers automáticos para 8h e 19h
 */
function criarTriggersAutomaticos() {
  // Remove triggers antigos primeiro
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'detectarFaturamento' ||
        trigger.getHandlerFunction() === 'limparFaturamentoDia') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Trigger às 8h (limpa e detecta)
  ScriptApp.newTrigger('limparFaturamentoDia')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  ScriptApp.newTrigger('detectarFaturamento')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();

  // Trigger às 19h (detecta)
  ScriptApp.newTrigger('detectarFaturamento')
    .timeBased()
    .atHour(19)
    .everyDays(1)
    .create();

  console.log('Triggers automáticos criados: 8h (limpa + detecta) e 19h (detecta)');

  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ Triggers automáticos configurados!\n\n' +
           '• 8h da manhã: Limpa faturamento anterior e inicia novo monitoramento\n' +
           '• 19h da noite: Detecta faturamento do dia\n\n' +
           'O sistema agora irá monitorar automaticamente.');
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
