var inputFile = document.getElementById("input");
var cmbEstados = document.getElementById("cmbEstados");
var codigos = [""];

inputFile.addEventListener("change", function() {
  if (codigos[0] != ""){
    readXlsxFile(inputFile.files[0], { getSheets: true }).then(function(sheets) {
      for (var i = 1; i <= sheets.length; i++) {
        readXlsxFile(inputFile.files[0], { sheet: i }).then(function(sheetData) {
          console.log(sheetData);
        })
      }
      console.log(sheets);
    })
  }
  else {
    alert("No se ha elegido un estado");
  }
  
})

cmbEstados.addEventListener("change", function() {
  switch(cmbEstados.value) {
    case "Aguascalientes":
      codigos[0] = "CAP-012-CMPC-2021";
      codigos.push("CON-046-LAFIWO-CEPC-2021");
      console.log(codigos)
      break;
      case "Colima":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "UEPC-CAP-855/19"
        console.log(codigos);
      break;
      case "Guanajuato":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "SSPG-220217-CEPC-278C"
        console.log(codigos)
      break;
      case "Jalisco":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "IEPC-060-07/2020"
        console.log(codigos)
      break;
      case "Michoacán":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "CEPCM-CONS-35-LFW-NJ1-2021"
        console.log(codigos)
      break;
      case "Zacatecas":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "CEPC/REG.PART/EM.CAP/CMPC/063/2020"
        console.log(codigos);
      break;
  }
})

generatePDF = () => {
  var doc = new jsPDF()

  doc.text('Hello world!', 100, 100)
  doc.save('a4.pdf')
}