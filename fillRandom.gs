function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('Show alert', 'showAlert')
      .addToUi();
}

// INICIALIZA VARIABLES
// =QUERY(CALCULOS!N1:AC14;"SELECT *") // HASTA CONSEGUIR QUE FUNCIONE
var positionAlgoSeo = 0 //Posición en la lista de Sheets de la hoja
var rangoValoresAlgoSeo = "N1:AC14"
var positionTablaFinalSeo = 1 //Posición en la lista de Sheets de la hoja
var rangoValoresTablaFinalSeo = "A4:P17"

// FUNCIONES
function recargar(){
  var cell = SpreadsheetApp.getActiveSheet().getRange("A4")
  var cellValue = "Memorizando..."
  cell.setValue(cellValue)
  
  // START CODIGO DIEGO
  vuelcaRango(positionAlgoSeo,rangoValoresAlgoSeo,positionTablaFinalSeo,rangoValoresTablaFinalSeo)

  var anf_value = "B17"
  var anf_array = "F22:F28"
  var anf_result_init = "N22"
  printResultOfRandomRange(anf_value,anf_array,anf_result_init)
  
  var af_value = "C17"
  var af_array = "F22:F28"
  var af_result_init = "O22"
  printResultOfRandomRange(af_value,af_array,af_result_init)
  
  var bnf_value = "D17"
  var bnf_array = "H22:H28"
  var bnf_result_init = "P22"
  printResultOfRandomRange(bnf_value,bnf_array,bnf_result_init)
  
  var bf_value = "E17"
  var bf_array = "H22:H28"
  var bf_result_init = "Q22"
  printResultOfRandomRange(bf_value,bf_array,bf_result_init)
  
  var cnf_value = "F17"
  var cnf_array = "J22:J25"
  var cnf_result_init = "R22"
  printResultOfRandomRange(cnf_value,cnf_array,cnf_result_init)
  
  var cf_value = "G17"
  var cf_array = "J22:J25"
  var cf_result_init = "S22"
  printResultOfRandomRange(cf_value,cf_array,cf_result_init)
  
  showAlert("Información botón","Ejecución acabada")
  // END CODIGO DIEGO
}
function vuelcaRango(sheetOrigen, rangoOrigen, sheetDestino, rangoDestino){
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetDestino].getRange(rangoDestino).clearContent()
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetDestino].getRange(rangoDestino).setValues(SpreadsheetApp.getActiveSpreadsheet().getSheets()[sheetOrigen].getRange(rangoOrigen).getValues())
}
function cargaRango(rango,valores){
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[positionTablaFinalSeo].getRange(rango).setValues(valores)
}
function memorizaRango(rango){
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[positionAlgoSeo].getRange(rango).getValues()
}

function getRandomInRange(range){
  return parseInt((Math.random() * range*10)%range)
  
}
function printResultOfRandomRange(contadorCell,rangoValores,resultadoCell){
  var num_value = SpreadsheetApp.getActiveSheet().getRange(contadorCell).getValue()
  var values_array = SpreadsheetApp.getActiveSheet().getRange(rangoValores)
  var cell_result_init = SpreadsheetApp.getActiveSheet().getRange(resultadoCell)
  clearSiguientesColumn(cell_result_init.getRow(),cell_result_init.getColumn())
  for(var i=0; i<num_value;i++){
    var tempCell = SpreadsheetApp.getActiveSheet().getRange(cell_result_init.getRow()+i, cell_result_init.getColumn())
    var tempValue = SpreadsheetApp.getActiveSheet().getRange(values_array.getRow() + getRandomInRange(values_array.getNumRows()), values_array.getColumn()).getValue()
    tempCell.setValue(tempValue)
  }
}
function clearSiguientesColumn(row,column){
  var cell_a_limpiar = SpreadsheetApp.getActiveSheet().getRange(row,column)
  if(cell_a_limpiar.getValue() != ""){
    cell_a_limpiar.setValue("")
    clearSiguientesColumn(row+1,column)
  }
}

function showAlert() {
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     'Información',
     'El botón ha acabado la ejecución.',
      ui.ButtonSet.OK);
}
