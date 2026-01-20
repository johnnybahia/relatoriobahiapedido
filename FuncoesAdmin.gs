/**
 * FUNÇÕES ADMINISTRATIVAS PARA DIAGNÓSTICO E CORREÇÃO
 * Use estas funções quando houver problemas com o faturamento
 */

/**
 * RESETAR SNAPSHOT - Use quando o faturamento estiver incorreto
 * Isso força o sistema a recalcular do zero na próxima verificação
 */
function resetarSnapshot() {
  try {
    var props = PropertiesService.getScriptProperties();

    Logger.log("🔄 Resetando snapshot do sistema de faturamento...");

    // Remove snapshot antigo
    props.deleteProperty('SNAPSHOT_DADOS1');
    props.deleteProperty('SNAPSHOT_TIMESTAMP');

    Logger.log("✅ Snapshot resetado com sucesso!");
    Logger.log("ℹ️  Na próxima execução do trigger, um novo snapshot será criado");
    Logger.log("⚠️  O faturamento acumulado do dia será preservado");

    return {
      sucesso: true,
      mensagem: "Snapshot resetado. Aguarde próxima verificação automática."
    };

  } catch (erro) {
    Logger.log("❌ Erro ao resetar snapshot: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * RESETAR FATURAMENTO DO DIA - Use para limpar o acumulado de hoje
 * CUIDADO: Isso apaga o faturamento detectado hoje!
 */
function resetarFaturamentoDia() {
  try {
    var props = PropertiesService.getScriptProperties();

    Logger.log("🔄 Resetando faturamento acumulado do dia...");

    // Remove faturamento acumulado
    props.deleteProperty('ULTIMO_FATURAMENTO');
    props.deleteProperty('ULTIMO_FATURAMENTO_TIMESTAMP');
    props.deleteProperty('FATURAMENTO_DATA');

    Logger.log("✅ Faturamento do dia resetado com sucesso!");
    Logger.log("⚠️  O snapshot NÃO foi alterado");

    return {
      sucesso: true,
      mensagem: "Faturamento do dia resetado."
    };

  } catch (erro) {
    Logger.log("❌ Erro ao resetar faturamento: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * VERIFICAR E CORRIGIR DATA DO FATURAMENTO ACUMULADO
 * Detecta se o faturamento acumulado é de outro dia e limpa automaticamente
 * Use esta função quando o card exibir faturamento de dias anteriores
 */
function verificarECorrigirDataFaturamento() {
  try {
    var props = PropertiesService.getScriptProperties();

    var dataAtual = new Date();
    var diaAtual = ("0" + dataAtual.getDate()).slice(-2) + "/" +
                   ("0" + (dataAtual.getMonth() + 1)).slice(-2) + "/" +
                   dataAtual.getFullYear();

    var diaArmazenado = props.getProperty('FATURAMENTO_DATA');
    var faturamento = props.getProperty('ULTIMO_FATURAMENTO');

    Logger.log("🔍 VERIFICANDO DATA DO FATURAMENTO ACUMULADO:");
    Logger.log("   Data atual: " + diaAtual);
    Logger.log("   Data armazenada: " + (diaArmazenado || "Nenhuma"));

    if (!diaArmazenado || !faturamento) {
      Logger.log("   ✅ Nenhum faturamento acumulado encontrado");
      return {
        sucesso: true,
        mensagem: "Nenhum faturamento acumulado",
        precisouLimpar: false
      };
    }

    if (diaArmazenado !== diaAtual) {
      Logger.log("   ⚠️ PROBLEMA DETECTADO!");
      Logger.log("   O faturamento acumulado é de outro dia (" + diaArmazenado + ")");
      Logger.log("   Isso faz o card exibir dados antigos como se fossem de hoje");
      Logger.log("\n   🔄 Limpando faturamento acumulado antigo...");

      props.deleteProperty('ULTIMO_FATURAMENTO');
      props.deleteProperty('ULTIMO_FATURAMENTO_TIMESTAMP');
      props.deleteProperty('FATURAMENTO_DATA');

      Logger.log("   ✅ Faturamento acumulado removido!");
      Logger.log("   ℹ️  Card agora exibirá o último registro do histórico");

      return {
        sucesso: true,
        mensagem: "Faturamento de " + diaArmazenado + " foi removido. Card atualizado.",
        precisouLimpar: true,
        dataRemovida: diaArmazenado
      };
    } else {
      Logger.log("   ✅ Data correta - faturamento é de hoje mesmo");
      return {
        sucesso: true,
        mensagem: "Data correta",
        precisouLimpar: false
      };
    }

  } catch (erro) {
    Logger.log("❌ Erro ao verificar data: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * DIAGNÓSTICO COMPLETO DO SISTEMA DE FATURAMENTO
 * Mostra o estado atual de todos os componentes
 */
function diagnosticarFaturamento() {
  try {
    Logger.log("\n" + "=".repeat(60));
    Logger.log("🔍 DIAGNÓSTICO DO SISTEMA DE FATURAMENTO");
    Logger.log("=".repeat(60));

    var props = PropertiesService.getScriptProperties();

    // 1. Snapshot
    Logger.log("\n📸 SNAPSHOT:");
    var snapshot = props.getProperty('SNAPSHOT_DADOS1');
    var snapshotTimestamp = props.getProperty('SNAPSHOT_TIMESTAMP');

    if (snapshot) {
      var mapaSnapshot = JSON.parse(snapshot);
      var totalOCs = Object.keys(mapaSnapshot).length;
      var valorTotalSnapshot = 0;

      Object.keys(mapaSnapshot).forEach(function(oc) {
        valorTotalSnapshot += mapaSnapshot[oc].valor;
      });

      Logger.log("   Status: ✅ Ativo");
      Logger.log("   Timestamp: " + snapshotTimestamp);
      Logger.log("   Total de OCs: " + totalOCs);
      Logger.log("   Valor total: R$ " + valorTotalSnapshot.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

      // Mostra as 5 maiores OCs
      var ocsOrdenadas = Object.keys(mapaSnapshot).map(function(oc) {
        return {
          oc: oc,
          cliente: mapaSnapshot[oc].cliente,
          valor: mapaSnapshot[oc].valor
        };
      }).sort(function(a, b) {
        return b.valor - a.valor;
      });

      Logger.log("\n   📊 Top 5 OCs por valor:");
      for (var i = 0; i < Math.min(5, ocsOrdenadas.length); i++) {
        var item = ocsOrdenadas[i];
        Logger.log("      " + (i+1) + ". OC " + item.oc + " - " + item.cliente + ": R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
      }

    } else {
      Logger.log("   Status: ❌ Não existe");
      Logger.log("   Ação: Execute o trigger manualmente ou aguarde próxima execução");
    }

    // 2. Dados Atuais (Dados1)
    Logger.log("\n📦 DADOS ATUAIS (DADOS1):");
    var dadosAtuais = lerDados1();

    if (dadosAtuais.length > 0) {
      var valorTotalAtual = 0;
      dadosAtuais.forEach(function(item) {
        valorTotalAtual += item.valor;
      });

      Logger.log("   Total de OCs: " + dadosAtuais.length);
      Logger.log("   Valor total: R$ " + valorTotalAtual.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

      // Agrupa por cliente
      var porCliente = {};
      dadosAtuais.forEach(function(item) {
        if (!porCliente[item.cliente]) {
          porCliente[item.cliente] = 0;
        }
        porCliente[item.cliente] += item.valor;
      });

      Logger.log("\n   📊 Top 5 clientes por valor:");
      var clientesOrdenados = Object.keys(porCliente).map(function(cliente) {
        return {
          cliente: cliente,
          valor: porCliente[cliente]
        };
      }).sort(function(a, b) {
        return b.valor - a.valor;
      });

      for (var i = 0; i < Math.min(5, clientesOrdenados.length); i++) {
        var item = clientesOrdenados[i];
        Logger.log("      " + (i+1) + ". " + item.cliente + ": R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
      }

    } else {
      Logger.log("   Status: ❌ Vazio");
    }

    // 3. Faturamento do Dia
    Logger.log("\n💰 FATURAMENTO ACUMULADO DO DIA:");
    var faturamento = props.getProperty('ULTIMO_FATURAMENTO');
    var faturamentoTimestamp = props.getProperty('ULTIMO_FATURAMENTO_TIMESTAMP');
    var faturamentoData = props.getProperty('FATURAMENTO_DATA');

    if (faturamento) {
      var dadosFaturamento = JSON.parse(faturamento);
      var totalFaturado = 0;

      dadosFaturamento.forEach(function(item) {
        totalFaturado += item.valor;
      });

      Logger.log("   Status: ✅ Ativo");
      Logger.log("   Data: " + faturamentoData);
      Logger.log("   Timestamp: " + faturamentoTimestamp);
      Logger.log("   Total de itens: " + dadosFaturamento.length);
      Logger.log("   Valor total faturado: R$ " + totalFaturado.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

      Logger.log("\n   📊 Detalhes:");
      dadosFaturamento.forEach(function(item) {
        Logger.log("      • " + item.cliente + " (" + item.marca + "): R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
      });

    } else {
      Logger.log("   Status: ❌ Nenhum faturamento detectado hoje");
    }

    // 4. Comparação Snapshot vs Atual
    if (snapshot && dadosAtuais.length > 0) {
      Logger.log("\n🔄 COMPARAÇÃO (SNAPSHOT vs ATUAL):");

      // Usa função agrupada para somar OCs repetidas
      var mapaAtual = agruparDados1PorOC();

      var mapaSnapshot = JSON.parse(snapshot);

      Logger.log("   OCs que sumiram (faturadas 100%):");
      var countSumiu = 0;
      Object.keys(mapaSnapshot).forEach(function(oc) {
        if (!mapaAtual[oc]) {
          countSumiu++;
          var item = mapaSnapshot[oc];
          if (countSumiu <= 10) { // Mostra apenas 10
            Logger.log("      • OC " + oc + " - " + item.cliente + ": R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
          }
        }
      });
      Logger.log("   Total: " + countSumiu + " OCs");

      Logger.log("\n   OCs com valor reduzido (faturamento parcial):");
      var countReduzido = 0;
      Object.keys(mapaSnapshot).forEach(function(oc) {
        var itemAnterior = mapaSnapshot[oc];
        var itemAtual = mapaAtual[oc];
        if (itemAtual && itemAtual.valor < itemAnterior.valor) {
          countReduzido++;
          var diferenca = itemAnterior.valor - itemAtual.valor;
          if (countReduzido <= 10) { // Mostra apenas 10
            Logger.log("      • OC " + oc + " - " + itemAnterior.cliente);
            Logger.log("        Era: R$ " + itemAnterior.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}) +
                      " | Agora: R$ " + itemAtual.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}) +
                      " | Faturado: R$ " + diferenca.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
          }
        }
      });
      Logger.log("   Total: " + countReduzido + " OCs");
    }

    Logger.log("\n" + "=".repeat(60));
    Logger.log("✅ Diagnóstico concluído!");
    Logger.log("=".repeat(60) + "\n");

    return {
      sucesso: true,
      mensagem: "Diagnóstico concluído. Verifique os logs."
    };

  } catch (erro) {
    Logger.log("❌ Erro no diagnóstico: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * FORÇAR RECÁLCULO COMPLETO
 * Reseta tudo e força nova detecção
 */
function forcarRecalculoCompleto() {
  try {
    Logger.log("🔄 Forçando recálculo completo do sistema...");

    // Reset snapshot
    resetarSnapshot();

    // Reset faturamento do dia
    resetarFaturamentoDia();

    Logger.log("✅ Sistema resetado completamente!");
    Logger.log("ℹ️  Execute 'getFaturamentoDia()' manualmente para criar novo snapshot");
    Logger.log("ℹ️  Ou aguarde a próxima execução automática do trigger");

    return {
      sucesso: true,
      mensagem: "Sistema resetado. Aguarde recálculo automático."
    };

  } catch (erro) {
    Logger.log("❌ Erro ao forçar recálculo: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * ANALISAR DETECÇÃO DE FATURAMENTO (DIAGNÓSTICO DETALHADO)
 * Mostra exatamente como o sistema está calculando o faturamento
 * Use esta função para rastrear de onde vêm os valores
 */
function analisarDeteccaoFaturamento() {
  try {
    Logger.log("═══════════════════════════════════════════════════════════");
    Logger.log("🔍 ANÁLISE DETALHADA DA DETECÇÃO DE FATURAMENTO");
    Logger.log("═══════════════════════════════════════════════════════════\n");

    var props = PropertiesService.getScriptProperties();
    var snapshotAnterior = props.getProperty('SNAPSHOT_DADOS1');
    var timestampSnapshot = props.getProperty('SNAPSHOT_TIMESTAMP');

    // 1. MOSTRA SNAPSHOT ANTERIOR
    Logger.log("📸 SNAPSHOT ANTERIOR (BASE DE COMPARAÇÃO):");
    Logger.log("   Timestamp: " + (timestampSnapshot || "Não disponível"));

    if (!snapshotAnterior) {
      Logger.log("   ❌ Nenhum snapshot encontrado!");
      Logger.log("\n⚠️ Criando snapshot inicial agora...");

      var mapaAtual = agruparDados1PorOC();
      props.setProperty('SNAPSHOT_DADOS1', JSON.stringify(mapaAtual));
      props.setProperty('SNAPSHOT_TIMESTAMP', obterTimestamp());

      Logger.log("✅ Snapshot criado com " + Object.keys(mapaAtual).length + " OCs");
      Logger.log("ℹ️  Execute esta função novamente após a próxima verificação automática");
      Logger.log("ℹ️  ou após alterações nos dados para ver as diferenças detectadas.\n");

      return {
        sucesso: true,
        mensagem: "Snapshot inicial criado. Execute novamente após próxima verificação."
      };
    }

    var mapaAnterior = JSON.parse(snapshotAnterior);
    var totalAnterior = 0;
    var countOCsAnterior = Object.keys(mapaAnterior).length;

    Logger.log("   Total de OCs: " + countOCsAnterior);
    Logger.log("\n   Detalhamento (primeiras 20 OCs):");

    var ocsAnterior = Object.keys(mapaAnterior).slice(0, 20);
    ocsAnterior.forEach(function(oc) {
      var item = mapaAnterior[oc];
      totalAnterior += item.valor;
      Logger.log("      • OC " + oc + " | " + item.cliente + " | R$ " +
                item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
    });

    if (countOCsAnterior > 20) {
      Logger.log("      ... e mais " + (countOCsAnterior - 20) + " OCs");
      Object.keys(mapaAnterior).slice(20).forEach(function(oc) {
        totalAnterior += mapaAnterior[oc].valor;
      });
    }

    Logger.log("\n   💰 VALOR TOTAL NO SNAPSHOT: R$ " + totalAnterior.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    // 2. MOSTRA ESTADO ATUAL
    Logger.log("\n" + "─".repeat(60));
    Logger.log("📊 ESTADO ATUAL (DADOS1 AGORA):");

    var mapaAtual = agruparDados1PorOC();
    var totalAtual = 0;
    var countOCsAtual = Object.keys(mapaAtual).length;

    Logger.log("   Total de OCs: " + countOCsAtual);
    Logger.log("\n   Detalhamento (primeiras 20 OCs):");

    var ocsAtual = Object.keys(mapaAtual).slice(0, 20);
    ocsAtual.forEach(function(oc) {
      var item = mapaAtual[oc];
      totalAtual += item.valor;
      Logger.log("      • OC " + oc + " | " + item.cliente + " | R$ " +
                item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
    });

    if (countOCsAtual > 20) {
      Logger.log("      ... e mais " + (countOCsAtual - 20) + " OCs");
      Object.keys(mapaAtual).slice(20).forEach(function(oc) {
        totalAtual += mapaAtual[oc].valor;
      });
    }

    Logger.log("\n   💰 VALOR TOTAL ATUAL: R$ " + totalAtual.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    // 3. COMPARAÇÃO DETALHADA
    Logger.log("\n" + "─".repeat(60));
    Logger.log("🔄 COMPARAÇÃO (O QUE MUDOU?):");

    var mapaOCMarca = criarMapaOCMarca();
    var faturado = [];
    var totalFaturado = 0;

    Logger.log("\n   OCs que SUMIRAM (faturadas 100%):");
    var countSumiu = 0;
    Object.keys(mapaAnterior).forEach(function(oc) {
      if (!mapaAtual[oc]) {
        countSumiu++;
        var item = mapaAnterior[oc];
        var marca = buscarMarcaNoMapa(oc, mapaOCMarca);

        if (countSumiu <= 15) {
          Logger.log("      • OC " + oc + " | " + item.cliente + " | " + marca + " | R$ " +
                    item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
        }

        faturado.push({
          cliente: item.cliente,
          marca: marca,
          valor: item.valor,
          oc: oc,
          tipo: "sumiu"
        });
        totalFaturado += item.valor;
      }
    });

    if (countSumiu === 0) {
      Logger.log("      (nenhuma)");
    } else if (countSumiu > 15) {
      Logger.log("      ... e mais " + (countSumiu - 15) + " OCs");
    }
    Logger.log("   Subtotal: R$ " + faturado.reduce(function(sum, item) {
      return item.tipo === "sumiu" ? sum + item.valor : sum;
    }, 0).toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    Logger.log("\n   OCs com VALOR REDUZIDO (faturamento parcial):");
    var countReduzido = 0;
    Object.keys(mapaAnterior).forEach(function(oc) {
      var itemAnterior = mapaAnterior[oc];
      var itemAtual = mapaAtual[oc];

      if (itemAtual && itemAtual.valor < itemAnterior.valor) {
        countReduzido++;
        var diferenca = itemAnterior.valor - itemAtual.valor;
        var marca = buscarMarcaNoMapa(oc, mapaOCMarca);

        if (countReduzido <= 15) {
          Logger.log("      • OC " + oc + " | " + itemAnterior.cliente + " | " + marca);
          Logger.log("        Era: R$ " + itemAnterior.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}) +
                    " → Agora: R$ " + itemAtual.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}) +
                    " → Faturado: R$ " + diferenca.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
        }

        faturado.push({
          cliente: itemAnterior.cliente,
          marca: marca,
          valor: diferenca,
          oc: oc,
          tipo: "reduzido"
        });
        totalFaturado += diferenca;
      }
    });

    if (countReduzido === 0) {
      Logger.log("      (nenhuma)");
    } else if (countReduzido > 15) {
      Logger.log("      ... e mais " + (countReduzido - 15) + " OCs");
    }

    Logger.log("\n   💰 TOTAL FATURADO DETECTADO: R$ " + totalFaturado.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    // 4. AGRUPAMENTO POR CLIENTE+MARCA
    Logger.log("\n" + "─".repeat(60));
    Logger.log("📦 AGRUPAMENTO POR CLIENTE+MARCA:");

    var faturadoAgrupado = {};
    faturado.forEach(function(item) {
      var chave = item.cliente + "|" + item.marca;
      if (!faturadoAgrupado[chave]) {
        faturadoAgrupado[chave] = {
          cliente: item.cliente,
          marca: item.marca,
          valor: 0,
          ocs: []
        };
      }
      faturadoAgrupado[chave].valor += item.valor;
      faturadoAgrupado[chave].ocs.push(item.oc);
    });

    var agrupados = Object.keys(faturadoAgrupado).map(function(chave) {
      return faturadoAgrupado[chave];
    }).sort(function(a, b) {
      return b.valor - a.valor;
    });

    Logger.log("\n   Total de grupos: " + agrupados.length);
    agrupados.forEach(function(item, index) {
      Logger.log("      " + (index + 1) + ". " + item.cliente + " | " + item.marca +
                " | R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
      Logger.log("         OCs: " + item.ocs.join(", "));
    });

    var totalAgrupado = agrupados.reduce(function(sum, item) { return sum + item.valor; }, 0);
    Logger.log("\n   💰 TOTAL AGRUPADO: R$ " + totalAgrupado.toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    // 5. O QUE ESTÁ NO HISTÓRICO
    Logger.log("\n" + "─".repeat(60));
    Logger.log("📋 O QUE ESTÁ NO HISTÓRICO (ABA HistoricoFaturamento):");

    var dataAtual = new Date();
    var diaAtual = ("0" + dataAtual.getDate()).slice(-2) + "/" +
                   ("0" + (dataAtual.getMonth() + 1)).slice(-2) + "/" +
                   dataAtual.getFullYear();

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HistoricoFaturamento");
    if (sheet && sheet.getLastRow() > 1) {
      var historicoDados = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
      var dadosHoje = [];

      historicoDados.forEach(function(row) {
        var dataRegistro = row[0];
        if (dataRegistro instanceof Date) {
          var d = dataRegistro;
          dataRegistro = ("0" + d.getDate()).slice(-2) + "/" +
                        ("0" + (d.getMonth() + 1)).slice(-2) + "/" +
                        d.getFullYear();
        } else {
          dataRegistro = dataRegistro.toString().trim();
        }

        if (dataRegistro === diaAtual) {
          dadosHoje.push({
            cliente: row[1].toString(),
            marca: row[2].toString(),
            valor: typeof row[3] === 'number' ? row[3] : parseFloat(row[3]) || 0,
            observacao: row[4] ? row[4].toString() : ""
          });
        }
      });

      Logger.log("   Data de hoje: " + diaAtual);
      Logger.log("   Registros de hoje: " + dadosHoje.length);

      if (dadosHoje.length > 0) {
        Logger.log("\n   Detalhamento:");
        var totalHistorico = 0;
        dadosHoje.forEach(function(item, index) {
          totalHistorico += item.valor;
          var obs = item.observacao ? " [" + item.observacao + "]" : "";
          Logger.log("      " + (index + 1) + ". " + item.cliente + " | " + item.marca +
                    " | R$ " + item.valor.toLocaleString('pt-BR', {minimumFractionDigits: 2}) + obs);
        });
        Logger.log("\n   💰 TOTAL NO HISTÓRICO HOJE: R$ " + totalHistorico.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
      } else {
        Logger.log("   (nenhum registro de hoje)");
      }
    } else {
      Logger.log("   ❌ Aba HistoricoFaturamento não encontrada ou vazia");
    }

    // 6. RESUMO FINAL
    Logger.log("\n" + "═".repeat(60));
    Logger.log("📊 RESUMO FINAL:");
    Logger.log("   Total anterior (snapshot): R$ " + totalAnterior.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
    Logger.log("   Total atual (Dados1): R$ " + totalAtual.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
    Logger.log("   Diferença detectada: R$ " + totalFaturado.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
    Logger.log("   Diferença real: R$ " + (totalAnterior - totalAtual).toLocaleString('pt-BR', {minimumFractionDigits: 2}));

    if (Math.abs(totalFaturado - (totalAnterior - totalAtual)) > 0.01) {
      Logger.log("\n   ⚠️ ATENÇÃO: Há divergência entre a diferença detectada e a real!");
      Logger.log("   Isso pode indicar problema no agrupamento ou na comparação.");
    }

    Logger.log("═".repeat(60));

    return {
      sucesso: true,
      totalAnterior: totalAnterior,
      totalAtual: totalAtual,
      totalFaturado: totalFaturado,
      diferencaReal: totalAnterior - totalAtual
    };

  } catch (erro) {
    Logger.log("❌ Erro ao analisar detecção: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * EXPORTAR DIAGNÓSTICO PARA PLANILHA
 * Cria uma aba com análise detalhada do faturamento
 * Facilita visualização dos dados de comparação
 */
function exportarDiagnosticoParaPlanilha() {
  try {
    Logger.log("📊 Exportando diagnóstico para planilha...");

    var doc = SpreadsheetApp.getActiveSpreadsheet();
    var nomeAba = "DEBUG_Faturamento";

    // Remove aba antiga se existir
    var abaExistente = doc.getSheetByName(nomeAba);
    if (abaExistente) {
      doc.deleteSheet(abaExistente);
    }

    // Cria nova aba
    var sheet = doc.insertSheet(nomeAba);

    var props = PropertiesService.getScriptProperties();
    var snapshotAnterior = props.getProperty('SNAPSHOT_DADOS1');
    var timestampSnapshot = props.getProperty('SNAPSHOT_TIMESTAMP');

    if (!snapshotAnterior) {
      sheet.appendRow(["⚠️ ERRO", "Nenhum snapshot encontrado!"]);
      sheet.appendRow([]);
      sheet.appendRow(["SOLUÇÃO:", "Execute a função getFaturamentoDia() primeiro para criar o snapshot inicial"]);
      sheet.appendRow(["OU", "Aguarde a próxima execução automática do trigger (a cada 1 hora)"]);

      Logger.log("⚠️ Nenhum snapshot encontrado. Criando um agora...");

      // Cria snapshot automaticamente
      var mapaAtual = agruparDados1PorOC();
      props.setProperty('SNAPSHOT_DADOS1', JSON.stringify(mapaAtual));
      props.setProperty('SNAPSHOT_TIMESTAMP', obterTimestamp());

      Logger.log("✅ Snapshot criado! Execute esta função novamente após próxima verificação.");

      return {
        sucesso: false,
        mensagem: "Snapshot não encontrado. Um novo foi criado. Execute novamente após próxima verificação."
      };
    }

    var mapaAnterior = JSON.parse(snapshotAnterior);
    var mapaAtual = agruparDados1PorOC();
    var mapaOCMarca = criarMapaOCMarca();

    // SEÇÃO 1: SNAPSHOT ANTERIOR
    sheet.appendRow(["📸 SNAPSHOT ANTERIOR", timestampSnapshot || ""]);
    sheet.appendRow([]);
    sheet.appendRow(["OC", "Cliente", "Valor"]);

    Object.keys(mapaAnterior).forEach(function(oc) {
      var item = mapaAnterior[oc];
      sheet.appendRow([oc, item.cliente, item.valor]);
    });

    var linhaAtual = sheet.getLastRow() + 2;

    // SEÇÃO 2: ESTADO ATUAL
    sheet.getRange(linhaAtual, 1).setValue("📊 ESTADO ATUAL");
    linhaAtual += 2;
    sheet.getRange(linhaAtual, 1, 1, 3).setValues([["OC", "Cliente", "Valor"]]);
    linhaAtual++;

    var linhaInicioAtual = linhaAtual;
    Object.keys(mapaAtual).forEach(function(oc) {
      var item = mapaAtual[oc];
      sheet.getRange(linhaAtual, 1, 1, 3).setValues([[oc, item.cliente, item.valor]]);
      linhaAtual++;
    });

    linhaAtual += 2;

    // SEÇÃO 3: DIFERENÇAS DETECTADAS
    sheet.getRange(linhaAtual, 1).setValue("🔄 FATURAMENTO DETECTADO");
    linhaAtual += 2;
    sheet.getRange(linhaAtual, 1, 1, 6).setValues([["OC", "Cliente", "Marca", "Valor Faturado", "Tipo", "Observação"]]);
    linhaAtual++;

    var linhaInicioFaturado = linhaAtual;

    // OCs que sumiram
    Object.keys(mapaAnterior).forEach(function(oc) {
      if (!mapaAtual[oc]) {
        var item = mapaAnterior[oc];
        var marca = buscarMarcaNoMapa(oc, mapaOCMarca);
        sheet.getRange(linhaAtual, 1, 1, 6).setValues([[
          oc,
          item.cliente,
          marca,
          item.valor,
          "Sumiu (100%)",
          "OC removida completamente"
        ]]);
        linhaAtual++;
      }
    });

    // OCs com valor reduzido
    Object.keys(mapaAnterior).forEach(function(oc) {
      var itemAnterior = mapaAnterior[oc];
      var itemAtual = mapaAtual[oc];

      if (itemAtual && itemAtual.valor < itemAnterior.valor) {
        var diferenca = itemAnterior.valor - itemAtual.valor;
        var marca = buscarMarcaNoMapa(oc, mapaOCMarca);
        sheet.getRange(linhaAtual, 1, 1, 6).setValues([[
          oc,
          itemAnterior.cliente,
          marca,
          diferenca,
          "Reduzido",
          "Era R$ " + itemAnterior.valor.toFixed(2) + " → Agora R$ " + itemAtual.valor.toFixed(2)
        ]]);
        linhaAtual++;
      }
    });

    // Formata cabeçalhos
    sheet.getRange(1, 1, 1, 3).setBackground("#4CAF50").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.getRange(3, 1, 1, 3).setBackground("#2196F3").setFontColor("#FFFFFF").setFontWeight("bold");

    var linhaHeader2 = linhaInicioAtual - 1;
    sheet.getRange(linhaHeader2, 1, 1, 3).setBackground("#2196F3").setFontColor("#FFFFFF").setFontWeight("bold");

    sheet.getRange(linhaInicioFaturado - 1, 1, 1, 6).setBackground("#FF9800").setFontColor("#FFFFFF").setFontWeight("bold");

    // Ajusta larguras
    sheet.setColumnWidth(1, 120); // OC
    sheet.setColumnWidth(2, 200); // Cliente
    sheet.setColumnWidth(3, 150); // Marca/Valor
    sheet.setColumnWidth(4, 120); // Valor Faturado
    sheet.setColumnWidth(5, 120); // Tipo
    sheet.setColumnWidth(6, 300); // Observação

    // Congela primeira linha
    sheet.setFrozenRows(1);

    Logger.log("✅ Diagnóstico exportado para aba '" + nomeAba + "'");
    Logger.log("ℹ️  Abra a planilha e veja a aba " + nomeAba + " para análise detalhada");

    return {
      sucesso: true,
      mensagem: "Diagnóstico exportado para aba '" + nomeAba + "'"
    };

  } catch (erro) {
    Logger.log("❌ Erro ao exportar diagnóstico: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}

/**
 * VERIFICAR INCONSISTÊNCIAS NOS DADOS
 * Analisa se existem OCs com múltiplos clientes diferentes
 * Use esta função para diagnosticar problemas de atribuição de faturamento
 */
function verificarInconsistenciasOCs() {
  try {
    Logger.log("🔍 VERIFICANDO INCONSISTÊNCIAS NOS DADOS DA ABA DADOS1...\n");

    var dados = lerDados1();
    var mapaOCs = {};
    var inconsistencias = [];

    // Agrupa todas as ocorrências de cada OC
    dados.forEach(function(item) {
      var oc = item.ordemCompra;

      if (!mapaOCs[oc]) {
        mapaOCs[oc] = {
          clientes: [item.cliente],
          valores: [item.valor],
          totalValor: item.valor
        };
      } else {
        mapaOCs[oc].valores.push(item.valor);
        mapaOCs[oc].totalValor += item.valor;

        // Verifica se cliente é diferente
        if (mapaOCs[oc].clientes.indexOf(item.cliente) === -1) {
          mapaOCs[oc].clientes.push(item.cliente);
        }
      }
    });

    // Identifica OCs com múltiplos clientes
    Object.keys(mapaOCs).forEach(function(oc) {
      var info = mapaOCs[oc];

      if (info.clientes.length > 1) {
        inconsistencias.push({
          oc: oc,
          clientes: info.clientes,
          valores: info.valores,
          totalValor: info.totalValor,
          qtdLinhas: info.valores.length
        });
      }
    });

    // Exibe resultados
    Logger.log("📊 RESUMO:");
    Logger.log("   Total de OCs analisadas: " + Object.keys(mapaOCs).length);
    Logger.log("   Total de linhas nos dados: " + dados.length);
    Logger.log("   OCs com múltiplos clientes: " + inconsistencias.length + "\n");

    if (inconsistencias.length > 0) {
      Logger.log("⚠️ INCONSISTÊNCIAS DETECTADAS:\n");

      inconsistencias.forEach(function(item, index) {
        Logger.log("   " + (index + 1) + ". OC: " + item.oc);
        Logger.log("      Clientes encontrados: " + item.clientes.join(", "));
        Logger.log("      Valores individuais: R$ " + item.valores.map(function(v) {
          return v.toLocaleString('pt-BR', {minimumFractionDigits: 2});
        }).join(", R$ "));
        Logger.log("      Total somado: R$ " + item.totalValor.toLocaleString('pt-BR', {minimumFractionDigits: 2}));
        Logger.log("      Quantidade de linhas: " + item.qtdLinhas);
        Logger.log("      ⚠️ PROBLEMA: Sistema manterá apenas '" + item.clientes[0] + "' (primeiro cliente)");
        Logger.log("");
      });

      Logger.log("❌ AÇÃO RECOMENDADA:");
      Logger.log("   Verifique a aba 'Dados1' e corrija os dados inconsistentes.");
      Logger.log("   Cada Ordem de Compra deveria pertencer a apenas um cliente.");

    } else {
      Logger.log("✅ NENHUMA INCONSISTÊNCIA DETECTADA!");
      Logger.log("   Todos os dados estão corretos: cada OC pertence a um único cliente.");
    }

    return {
      sucesso: true,
      totalOCs: Object.keys(mapaOCs).length,
      totalLinhas: dados.length,
      inconsistencias: inconsistencias.length
    };

  } catch (erro) {
    Logger.log("❌ Erro ao verificar inconsistências: " + erro.toString());
    return {
      sucesso: false,
      mensagem: "Erro: " + erro.toString()
    };
  }
}
