function onEdit(e){
    // Melhoria: Utilize o evento para verificar a célula editada diretamente
    if (e.range.getA1Notation() === "H6" && e.source.getSheetName() === "EFETIVO") {
        Pesquisa();
    }
}

function Pesquisa() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EFETIVO");
  const criterio = sheet.getRange("H6").getValue();

  console.log("Critério de pesquisa: " + criterio); // Usando console.log (Logger.log é antigo)

  if (criterio == 0){
    console.log("Critério de pesquisa é zero, função encerrada sem ação");
    return;
  }

  sheet.getRange("B6:G6").clearContent();
  console.log("Conteúdo de B6 até G6 limpo");

  // Pre-cache last row index and use a more direct approach
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(9, 2, lastRow - 8, 6);
  const dados = dataRange.getValues();
  console.log("Dados carregados para pesquisa: " + dados.length + " linhas encontradas");

  const filteredData = dados.filter(row => row[1] === criterio);
  console.log("Número de correspondências encontradas: " + filteredData.length);

  if (filteredData.length > 0) {
    sheet.getRange(6, 2, filteredData.length, 6).setValues(filteredData);
    console.log("Dados correspondentes inseridos de volta na planilha");
  } else {
    SpreadsheetApp.getUi().alert("Não encontrado registros para esta pesquisa!");
    console.log("Não foram encontrados registros para a pesquisa");
  }
}
