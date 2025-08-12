# Tratamento-estoque

Este script do Google Sheets automatiza o tratamento de dados brutos, removendo linhas em branco e organizando os dados importantes. Após o processamento, os dados são colados em uma aba de destino, limpando os dados anteriores, mas **mantendo a formatação e as fórmulas das outras colunas**.

## Funcionalidades

- Lê dados da aba "TRATAMENTO_ESTOQUE_SISREDE"
- Filtra linhas que possuem código (não vazias)
- Seleciona colunas específicas (Código, Descrição, Consumo, Posição)
- Cola os dados tratados na aba "ESTOQUE SIS REDE PREGÃO"
- Limpa o conteúdo anterior das colunas de destino, sem apagar formatação ou fórmulas existentes
- Atualiza automaticamente a fórmula da coluna "E" para as novas linhas adicionadas

## Como usar

1. Copie o código do script para o editor de script do Google Sheets.
2. Ajuste os nomes das abas, se necessário.
3. Execute a função `tratarDados()` para realizar o processamento.
4. Verifique os resultados na aba "ESTOQUE SIS REDE PREGÃO".

## Requisitos

- Planilha Google Sheets com as abas:
  - `TRATAMENTO_ESTOQUE_SISREDE` com os dados brutos na estrutura esperada
  - `ESTOQUE SIS REDE PREGÃO` formatada com fórmulas na coluna E (a fórmula da linha 2 será replicada automaticamente)

## Código do script

```javascript
function tratarDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBruta = ss.getSheetByName("TRATAMENTO_ESTOQUE_SISREDE");
  const abaDestino = ss.getSheetByName("ESTOQUE SIS REDE PREGÃO");
  
  if (!abaDestino) {
    throw new Error('Aba "ESTOQUE SIS REDE PREGÃO" não encontrada!');
  }
  
  const dados = abaBruta.getDataRange().getValues();
  const dadosFiltrados = dados
    .slice(5)
    .filter(linha => linha[0] && linha[0].toString().trim() !== "");
  
  const tabelaFinal = [];
  tabelaFinal.push(["Código", "Descrição", "Consumo", "Posição"]);
  
  dadosFiltrados.forEach(linha => {
    const codigo = linha[0];
    const descricao = linha[1];
    const consumo = linha[8];
    const posicao = linha[9];
    tabelaFinal.push([codigo, descricao, consumo, posicao]);
  });
  
  const numLinhasParaLimpar = abaDestino.getLastRow();
  if(numLinhasParaLimpar > 0) {
    abaDestino.getRange(1, 1, numLinhasParaLimpar, 4).clearContent();
  }
  
  abaDestino.getRange(1, 1, tabelaFinal.length, tabelaFinal[0].length).setValues(tabelaFinal);
  
  if(tabelaFinal.length > 1) {
    const numLinhasFormula = tabelaFinal.length - 1;
    const celulaModelo = abaDestino.getRange(2, 5);
    const formulaModelo = celulaModelo.getFormula();
    
    if (formulaModelo) {
      const rangeParaFormulas = abaDestino.getRange(3, 5, numLinhasFormula - 1, 1);
      rangeParaFormulas.setFormula(formulaModelo);
    }
  }
}
