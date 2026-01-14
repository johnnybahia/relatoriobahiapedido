# Sistema de Faturamento - Instruções de Instalação

## 📋 O que foi implementado

### Backend (Google Apps Script)
✅ Funções para ler aba "Dados1" (Ordem de Compra, Valor, Cliente)
✅ Sistema de snapshot para comparação de dados
✅ Detecção automática de faturamento
✅ Agregação por cliente e marca
✅ Triggers automáticos às 8h e 19h
✅ Funções manuais acessíveis pelo menu

### Frontend (Webapp)
✅ Card "Pedidos a Faturar" - mostra valores atuais agrupados
✅ Card "Faturamento do Dia" - mostra o que foi faturado
✅ Atualização automática a cada 5 minutos
✅ Design responsivo e moderno

---

## 🚀 Como instalar

### 1. Atualizar o código do Google Apps Script

O arquivo `macros.gs` já foi atualizado com todas as funções necessárias. Basta copiar o conteúdo para o Google Apps Script Editor.

### 2. Adicionar os cards no Webapp (Index.html)

1. Abra o Google Apps Script Editor
2. Localize o arquivo **Index.html** na lista de arquivos
3. Encontre onde está o card "Totais por Marca"
4. Logo ABAIXO desse card, adicione o conteúdo do arquivo `cards_faturamento.html`
5. Salve o arquivo

**Estrutura esperada no Index.html:**
```html
<!-- ... código existente ... -->

<!-- Card Totais por Marca (já existe) -->
<div class="card" id="cardTotaisMarca">
  ...
</div>

<!-- ADICIONE AQUI OS NOVOS CARDS -->
<!-- Copie todo o conteúdo de cards_faturamento.html -->

<!-- ... resto do código ... -->
```

### 3. Configurar Triggers Automáticos

No Google Sheets, vá em:
1. **Menu:** `🏭 Relatórios` → `💰 Faturamento` → `⚙️ Configurar Triggers Automáticos`
2. Clique e aguarde a confirmação

Isso criará triggers para:
- **8h da manhã:** Limpa faturamento anterior e inicia novo monitoramento
- **19h da noite:** Detecta faturamento do dia

### 4. Primeira Execução (Importante!)

Antes de usar, execute UMA VEZ:
1. **Menu:** `🏭 Relatórios` → `💰 Faturamento` → `🔄 Atualizar Faturamento Agora`

Isso cria o primeiro snapshot dos dados.

---

## 💡 Como funciona

### Lógica de Faturamento

O sistema compara os dados da aba "Dados1" em dois momentos diferentes:

| Situação | O que significa | Resultado |
|----------|----------------|-----------|
| **OC sumiu** | Pedido foi totalmente faturado | Valor total vai para "Faturamento do Dia" |
| **Valor diminuiu** | Faturamento parcial | Diferença vai para "Faturamento do Dia" |
| **Valor aumentou** | Novo pedido ou acréscimo | Fica em "Pedidos a Faturar" |
| **OC nova** | Pedido novo | Fica em "Pedidos a Faturar" |

### Exemplo Prático

**8h da manhã:**
```
OC 12345 | Cliente: João Silva | Valor: R$ 10.000
OC 67890 | Cliente: Maria Costa | Valor: R$ 5.000
```

**19h da noite:**
```
OC 12345 | Cliente: João Silva | Valor: R$ 3.000 (diminuiu R$ 7.000)
OC 11111 | Cliente: Pedro Lima | Valor: R$ 8.000 (nova OC)
```

**Resultado:**
- **Faturamento do Dia:**
  - João Silva: R$ 7.000 (faturamento parcial)
  - Maria Costa: R$ 5.000 (OC sumiu = faturou tudo)
  - **Total: R$ 12.000**

- **Pedidos a Faturar:**
  - João Silva: R$ 3.000 (restante)
  - Pedro Lima: R$ 8.000 (novo)
  - **Total: R$ 11.000**

---

## 🎯 Funções Disponíveis no Menu

### `🏭 Relatórios` → `💰 Faturamento`

| Função | O que faz |
|--------|-----------|
| **🔄 Atualizar Faturamento Agora** | Executa detecção manualmente (use quando quiser verificar) |
| **⚙️ Configurar Triggers Automáticos** | Cria triggers para 8h e 19h (executar apenas 1 vez) |
| **🧹 Limpar Faturamento do Dia** | Zera faturamento (use no início do dia se necessário) |

---

## 📊 Cards no Webapp

### Card 1: Pedidos a Faturar 💼
- Mostra todos os pedidos atuais na aba "Dados1"
- Agrupa por cliente e marca
- Soma valores de cada cliente/marca
- Atualiza automaticamente
- **Timestamp:** Mostra quando foi atualizado

### Card 2: Faturamento do Dia 💰
- Mostra o que foi faturado desde às 8h
- Detecta OCs que sumiram ou valores que diminuíram
- Agrupa por cliente e marca
- **Timestamp:** Mostra quando foi a última verificação (8h ou 19h)

---

## 🔧 Troubleshooting

### Os cards não aparecem no webapp
- Verifique se copiou o HTML para o Index.html
- Verifique se está ABAIXO do card "Totais por Marca"
- Limpe o cache do navegador (Ctrl+Shift+R)

### Faturamento não está sendo detectado
- Execute manualmente: Menu → Faturamento → Atualizar Faturamento Agora
- Verifique se a aba "Dados1" existe e tem as colunas corretas
- Verifique se os triggers foram criados: Menu → Configurar Triggers Automáticos

### Triggers não estão funcionando
- Vá em Apps Script Editor → Triggers (ícone de relógio)
- Verifique se existem triggers para "detectarFaturamento" e "limparFaturamentoDia"
- Se não existirem, execute: Menu → Configurar Triggers Automáticos

### Valores errados ou marca "N/A"
- Verifique se a aba "MARCAS" tem a coluna "ORDEM DE COMPRA" preenchida
- Verifique se os valores na aba "Dados1" estão corretos (coluna VALOR deve ser número)

---

## 📝 Estrutura da Aba "Dados1"

A aba deve ter exatamente estas 3 colunas na linha 1:

| Coluna A | Coluna B | Coluna C |
|----------|----------|----------|
| ORDEM DE COMPRA | VALOR | CLIENTE |

**Exemplo de dados:**
```
ORDEM DE COMPRA | VALOR     | CLIENTE
12345          | 10000.00  | João Silva
67890          | 5000.50   | Maria Costa
11223          | 8000.00   | Pedro Lima
```

---

## ✅ Checklist de Instalação

- [ ] Código `macros.gs` atualizado no Apps Script
- [ ] HTML dos cards adicionado no `Index.html`
- [ ] Triggers automáticos configurados (Menu → Configurar Triggers)
- [ ] Primeira execução manual realizada (Menu → Atualizar Faturamento Agora)
- [ ] Webapp atualizado (recarregue a página)
- [ ] Aba "Dados1" criada com as 3 colunas corretas
- [ ] Aba "MARCAS" tem as OCs cadastradas

---

## 🎉 Pronto!

Após seguir todos os passos, o sistema estará funcionando:
- ✅ Cards aparecendo no webapp
- ✅ Atualização automática às 8h e 19h
- ✅ Detecção de faturamento funcionando
- ✅ Menu com opções manuais disponível

**Qualquer dúvida, execute a função de debug no menu e verifique os logs!**
