function onEdit(){

    /* Esse código faz com que execute a função Pesquisa() somente quando estivermos manipulando a celular H6 */
    var guiaMenu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EFETIVO");
    var celula = guiaMenu.getActiveCell().getA1Notation();

    if(celula == "H6"){
        Pesquisa();
      }
}

function Pesquisa() {
  var guiaDados = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EFETIVO");
  var criterio = guiaDados.getRange("H6").getValue();

  Logger.log("Critério de pesquisa: " + criterio); // Imprime o critério de pesquisa

  guiaDados.getRange("B6:G6").clearContent();
  Logger.log("Conteúdo de B6 até G6 limpo"); // Confirma a limpeza das células

  if (criterio == 0){
    Logger.log("Critério de pesquisa é zero, função encerrada sem ação");
    return;
  }

  var dados = guiaDados.getRange(9,2,guiaDados.getLastRow() - 8, 6).getValues();
  Logger.log("Dados carregados para pesquisa: " + dados.length + " linhas encontradas"); // Informa quantas linhas de dados foram carregadas

  var pesquisa = dados.filter(function(value, i, arr) { return criterio == arr[i][1]; });
  Logger.log("Número de correspondências encontradas: " + pesquisa); // Mostra quantos resultados correspondem ao critério

  if (pesquisa.length > 0) {
    guiaDados.getRange(6,2, pesquisa.length, pesquisa[0].length).setValues(pesquisa);
    Logger.log("Dados correspondentes inseridos de volta na planilha");
  } else {
    Browser.msgBox("Não encontrado registros para esta pesquisa!");
    Logger.log("Não foram encontrados registros para a pesquisa"); // Mensagem de alerta se nenhum registro foi encontrado
  }
}
