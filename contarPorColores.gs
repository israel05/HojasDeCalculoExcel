function contarColorForma() {
  const libro = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = libro.getActiveSheet()
  const colorRojo = hoja.getRange('I1').getBackground();
  const colorNaranja = hoja.getRange('I2').getBackground();
  const colorVerde = hoja.getRange('I3').getBackground();
  var contadorRojo = 0;
  var contadorNaranja = 0;
  var contadorVerde = 0;
  for (var i= 8; i<=12;i++){
      for (var j= 4; j<=6; j++){
  
            const celdaActual = hoja.getRange(i,j)
            if(celdaActual.getBackground() == colorRojo){
              contadorRojo++
            }
            if(celdaActual.getBackground() == colorNaranja){
              contadorNaranja++
            }
            if(celdaActual.getBackground() == colorVerde){
              contadorVerde++
            }
      }

  }
 hoja.getRange('G9').setValue(contadorRojo)
  hoja.getRange('G10').setValue(contadorNaranja)
  hoja.getRange('G11').setValue(contadorVerde)

}

function contarColorExpresion() {
  const libro = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = libro.getActiveSheet()
  const colorRojo = hoja.getRange('I1').getBackground();
  const colorNaranja = hoja.getRange('I2').getBackground();
  const colorVerde = hoja.getRange('I3').getBackground();
  var contadorRojo = 0;
  var contadorNaranja = 0;
  var contadorVerde = 0;
  for (var i= 13; i<=19;i++){
      for (var j= 4; j<=6; j++){
  
            const celdaActual = hoja.getRange(i,j)
            if(celdaActual.getBackground() == colorRojo){
              contadorRojo++
            }
            if(celdaActual.getBackground() == colorNaranja){
              contadorNaranja++
            }
            if(celdaActual.getBackground() == colorVerde){
              contadorVerde++
            }
      }

  }
 hoja.getRange('G15').setValue(contadorRojo)
  hoja.getRange('G16').setValue(contadorNaranja)
  hoja.getRange('G17').setValue(contadorVerde)

}

function contarColorFondo() {
  const libro = SpreadsheetApp.getActiveSpreadsheet()
  const hoja = libro.getActiveSheet()
  const colorRojo = hoja.getRange('I1').getBackground();
  const colorNaranja = hoja.getRange('I2').getBackground();
  const colorVerde = hoja.getRange('I3').getBackground();
  var contadorRojo = 0;
  var contadorNaranja = 0;
  var contadorVerde = 0;
  for (var i= 20; i<=24;i++){
      for (var j= 4; j<=6; j++){
  
            const celdaActual = hoja.getRange(i,j)
            if(celdaActual.getBackground() == colorRojo){
              contadorRojo++
            }
            if(celdaActual.getBackground() == colorNaranja){
              contadorNaranja++
            }
            if(celdaActual.getBackground() == colorVerde){
              contadorVerde++
            }
      }

  }
 hoja.getRange('G21').setValue(contadorRojo)
  hoja.getRange('G22').setValue(contadorNaranja)
  hoja.getRange('G23').setValue(contadorVerde)

}

function contarColor(){
  contarColorForma();
  contarColorExpresion();
  contarColorFondo();
}

function onOpen(){
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("Color");
  menu.addItem("cuentacolores","contarColor")
  menu.addToUi();
}
