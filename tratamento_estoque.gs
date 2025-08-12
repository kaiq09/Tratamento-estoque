function tratarDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaBruta = ss.getSheetByName("TRATAMENTO_ESTOQUE_SISREDE");
  const abaDestino = ss.getSheetByName("ESTOQUE SIS REDE PREGÃO");
  
  if (!abaDestino) {
    throw new Error('Aba "ESTOQUE SIS REDE PREGÃO" não encontrada!');
  }
  
  // Lê todos os dados da aba bruta
  const dados = abaBruta.getDataRange().getValues();
  
  // Filtra a partir da linha 6, ignorando linhas sem código
  const dadosFiltrados = dados
    .slice(5)
    .filter(linha => linha[0] && linha[0].toString().trim() !== "");
  
  // Monta tabela com as colunas que quer: Código (A), Descrição (B), Consumo (C), Posição (D)
  const tabelaFinal = [];
  
  // Cabeçalho (A, B, C, D)
  tabelaFinal.push(["Código", "Descrição", "Consumo", "Posição"]);
  
  dadosFiltrados.forEach(linha => {
    const codigo = linha[0];
    const descricao = linha[1];
    const consumo = linha[8];
    const posicao = linha[9];
    tabelaFinal.push([codigo, descricao, consumo, posicao]);
  });
  
  // Limpa só as colunas A-D da aba destino (para evitar restos)
  const numLinhasParaLimpar = abaDestino.getLastRow();
  if(numLinhasParaLimpar > 0) {
    abaDestino.getRange(1, 1, numLinhasParaLimpar, 4).clearContent();
  }
  
  // Escreve os dados nas colunas A-D, começando da linha 1
  abaDestino.getRange(1, 1, tabelaFinal.length, tabelaFinal[0].length).setValues(tabelaFinal);
  
  // Agora copia a fórmula da coluna E da linha 2 para as demais linhas que têm dados
  
  // Se a tabelaFinal tem mais que só o cabeçalho
  if(tabelaFinal.length > 1) {
    const numLinhasFormula = tabelaFinal.length - 1; // desconta o cabeçalho
    // Pega a fórmula da célula E2
    const celulaModelo = abaDestino.getRange(2, 5);
    const formulaModelo = celulaModelo.getFormula();
    
    if (formulaModelo) {
      // Copia a fórmula para as linhas da 3 até a última com dados
      const rangeParaFormulas = abaDestino.getRange(3, 5, numLinhasFormula - 1, 1);
      // Preenche esse range com a mesma fórmula (o Google Sheets ajusta referências automaticamente)
      rangeParaFormulas.setFormula(formulaModelo);
    }
  }
}

