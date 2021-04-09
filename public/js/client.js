var inputFile = document.getElementById("input");
var borde = document.getElementById("borde");
var esquina = document.getElementById("esquina");
var color01 = document.getElementById("color01");
var color02 = document.getElementById("color02");
var color03 = document.getElementById("color03");
var color04 = document.getElementById("color04");
var color05 = document.getElementById("color05");
var color06 = document.getElementById("color06");
var color07 = document.getElementById("color07");
var color08 = document.getElementById("color08");
var color09 = document.getElementById("color09");
var color10 = document.getElementById("color10");
var txtFile = document.getElementById("txtFile");
var cmbEstados = document.getElementById("cmbEstados");
var codigos = [""];


color01.addEventListener("mouseover", function() {
  color01.style.zIndex = 50;
})

color01.addEventListener("mouseout", function() {
  color01.style.zIndex = 49;
})

color02.addEventListener("mouseover", function() {
  color02.style.zIndex = 50;
})

color02.addEventListener("mouseout", function() {
  color01.style.zIndex = 49;
})

color03.addEventListener("mouseover", function() {
  color03.style.zIndex = 50;
})

color03.addEventListener("mouseout", function() {
  color03.style.zIndex = 49;
})

color03.addEventListener("click", function() {
  borde.src = "../img/borde/borde-verde-claro.png";
  esquina.src = "../img/esquina/esquina-verde-claro.png"
})

color04.addEventListener("mouseover", function() {
  color04.style.zIndex = 50;
})

color04.addEventListener("mouseout", function() {
  color04.style.zIndex = 49;
})

color04.addEventListener("click", function() {
  borde.src = "../img/borde/borde-amarillo.png";
  esquina.src = "../img/esquina/esquina-amarillo.png"
})

color05.addEventListener("mouseover", function() {
  color05.style.zIndex = 50;
})

color05.addEventListener("mouseout", function() {
  color05.style.zIndex = 49;
})

color05.addEventListener("click", function() {
  borde.src = "../img/borde/borde-naranja.PNG";
  esquina.src = "../img/esquina/esquina-naranja.PNG"
})

color06.addEventListener("mouseover", function() {
  color06.style.zIndex = 50;
})

color06.addEventListener("mouseout", function() {
  color06.style.zIndex = 49;
})

color06.addEventListener("click", function() {
  borde.src = "../img/borde/borde-rojo.png";
  esquina.src = "../img/esquina/esquina-rojo.png"
})

color07.addEventListener("mouseover", function() {
  color07.style.zIndex = 50;
})

color07.addEventListener("mouseout", function() {
  color07.style.zIndex = 49;
})

color08.addEventListener("mouseover", function() {
  color08.style.zIndex = 50;
})

color08.addEventListener("mouseout", function() {
  color08.style.zIndex = 49;
})

color09.addEventListener("mouseover", function() {
  color09.style.zIndex = 50;
})

color09.addEventListener("mouseout", function() {
  color09.style.zIndex = 49;
})

color10.addEventListener("mouseover", function() {
  color10.style.zIndex = 50;
})

color10.addEventListener("mouseout", function() {
  color10.style.zIndex = 49;
})

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
    txtFile.innerHTML = inputFile.files[0].name;
    txtFile.style.height = "auto"
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
  var doc = new jsPDF('l', 'cm', [19.05, 25.4])
  doc.addImage(document.getElementById("esquina"), "PNG", 18, 0, 8, 8);
  doc.addImage(document.getElementById("borde"), "PNG", 0, 17);
  doc.addPage(25.4, 19.05)
  doc.text('Hello world!', 1, 1)
  doc.save('a4.pdf')
}