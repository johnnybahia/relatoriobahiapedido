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
