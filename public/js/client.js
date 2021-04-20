import { font_normal } from '../css/calibril-normal.js';
import { font_bold } from '../css/calibril-bold.js';
var inputFile = document.getElementById("input");
var inputLogo = document.getElementById("inputLogo");
var imgLogo = document.getElementById("imgLogo");
var imgLavi = new Image()
imgLavi.src = "./img/LAVI Fire_HI-RES.jpg"
var borde = document.getElementById("borde");
var loader = document.getElementById("loader");
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
var txtCity = document.getElementById("txtCity");
var cmbEstados = document.getElementById("cmbEstados");
var codigos = [""];
var tipo = "";
var txtInfo = "";
var text = "";
var image = new Image()
image.src = "./img/LAVI Fire Profile Pic.png"
var finished = 0;
var fraccion = 0;
var actual = 0;
var calibri_light_b64;



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
  borde.src = "./img/borde/borde-verde-claro.png";
  esquina.src = "./img/esquina/esquina-verde-claro.png"
})

color04.addEventListener("mouseover", function() {
  color04.style.zIndex = 50;
})

color04.addEventListener("mouseout", function() {
  color04.style.zIndex = 49;
})

color04.addEventListener("click", function() {
  borde.src = "./img/borde/borde-amarillo.png";
  esquina.src = "./img/esquina/esquina-amarillo.png"
})

color05.addEventListener("mouseover", function() {
  color05.style.zIndex = 50;
})

color05.addEventListener("mouseout", function() {
  color05.style.zIndex = 49;
})

color05.addEventListener("click", function() {
  borde.src = "./img/borde/borde-naranja.png";
  esquina.src = "./img/esquina/esquina-naranja.png"
})

color06.addEventListener("mouseover", function() {
  color06.style.zIndex = 50;
})

color06.addEventListener("mouseout", function() {
  color06.style.zIndex = 49;
})

color06.addEventListener("click", function() {
  borde.src = "./img/borde/borde-rojo.png";
  esquina.src = "./img/esquina/esquina-rojo.png"
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

inputFile.addEventListener("change", async function() {
  if (codigos[0] != ""){
    console.log(inputFile.files[0])
    loader.style.display = "block"
    readXlsxFile(inputFile.files[0], { getSheets: true }).then(function(sheets) {
      fraccion = 100 / sheets.length
      for (var i = 1; i <= sheets.length; i++) {
        readXlsxFile(inputFile.files[0], { sheet: i }).then(function(sheetData) {
          var date = sheetData[3];
          var fecha = new Date(Date.parse(date[1]))
          var mes = "";
          var mesNumber = "";
          switch (fecha.getMonth())
          {
            case 0:
              mes = "Enero";
              mesNumber = "01";
            break;
            case 1:
              mes = "Febrero";
              mesNumber = "02";
            break;
            case 2:
              mes = "Marzo";
              mesNumber = "03";
            break;
            case 3:
              mes = "Abril";
              mesNumber = "04";
            break;
            case 4:
              mes = "Mayo";
              mesNumber = "05";
            break;
            case 5:
              mes = "Junio";
              mesNumber = "06";
            break;
            case 6:
              mes = "Julio";
              mesNumber = "07";
            break;
            case 7:
              mes = "Agosto";
              mesNumber = "08";
            break;
            case 8:
              mes = "Septiembre";
              mesNumber = "09";
            break;
            case 9:
              mes = "Octubre";
              mesNumber = "10";
            break;
            case 10:
              mes = "Noviembre";
              mesNumber = "11";
            break;
            case 11:
              mes = "Diciembre";
              mesNumber = "12";
            break;
          }
          var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
          var ciudadFecha = txtCity.value + ", " + cmbEstados.value + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
          var razonSocial = sheetData[0]
          txtInfo = 'Impartido el día **'+ dia + " de " + mes + " de " + fecha.getFullYear() +'**, para personal de **'+ razonSocial[1]
          text = 'Impartido el día '+ dia + " de " + mes + " de " + fecha.getFullYear() +', para personal de '+ razonSocial[1]
          var nombreComercial = sheetData[1][1]
          var nombreCurso = sheetData[2][1]
          var founded = true;
          var position = 0;
          do
          {
            
            position = nombreCurso.indexOf("/")
            
            if (position != -1) 
            {
              
              nombreCurso = nombreCurso.replace('/', '-')
            }
            else 
            {
              founded = false;
            }
          }
          while (founded == true);
          var filename = dia + mesNumber + fecha.getFullYear().toString().substr(-2) + "_Constancias - " + nombreCurso + " - " + nombreComercial
          // .getDate devuelve el dia
          // .getMonth devuelve el mes 0= Enero, 1 = Febrero, etc
          // .getFullYear devuelve el año completo ej: 1995
          generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length);
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
      tipo = "Capacitador en Materia de Protección Civil";
      console.log(codigos)
      break;
      case "Colima":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "UEPC-CAP-855/19"
        tipo = "Capacitador en Materia de Protección Civil";
        console.log(codigos);
      break;
      case "Guanajuato":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "SSPG-220217-CEPC-278C"
        tipo = "Consultor y Capacitador";
        console.log(codigos)
      break;
      case "Jalisco":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "IEPC-060-07/2020"
        tipo = "Capacitador en Materia de Protección Civil";
        console.log(codigos)
      break;
      case "Michoacán":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "CEPCM-CONS-35-LFW-NJ1-2021"
        tipo = "Consultor y Capacitador";
        console.log(codigos)
      break;
      case "Zacatecas":
        if(codigos.length > 1) {
          codigos.pop();
        }
        codigos[0] = "CEPC/REG.PART/EM.CAP/CMPC/063/2020"
        tipo = "Consultor y Capacitador";
        console.log(codigos);
      break;
  }
})

var generatePDF = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength) => {

  for (var a = 5; a < sheetData.length; a++)
  {
    if (a == 5)
    {
      var doc = new jsPDF('l', 'cm', [19.05, 25.4])
      console.log("Entro")
      doc.addFileToVFS("calibril-normal.ttf", font_normal)
      doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
      doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
      doc.addFileToVFS("calibril-bold.ttf", font_bold)
      doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
      doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
    }
    else
    {
      doc.addPage(25.4, 19.05)
    }
    doc.setFont("calibri-normal");
    doc.setTextColor( 0, 0, 0 )
    doc.addImage(esquina, "PNG", 20.4, 0, 5, 5);
    doc.addImage(borde, "PNG", 0, 17.5, 30, 1.5);
    doc.addImage(imgLavi, "JPEG", 0.5, 0.5, 9, 2.7); // , 9, 2.5
    
    doc.addImage(image, "png", 2, 15.7, 1.8, 1.8)

    doc.setFontSize(10);
    doc.text(ciudadFecha, 16, 3);
    doc.text(ciudadFecha, 23, 4);
    var xposition = (doc.getStringUnitWidth(ciudadFecha) / 800 ) / 2
    doc.text("LAVI Fire Workshop México, S.A. de C.V.", 25.4/2, 11, 'center');
    doc.text(tipo, 25.4/2, 11.5, 'center');
    doc.text(codigos[0], 25.4/2, 12, 'center');
    if (codigos.length > 1)
    {
      doc.text(codigos[1], 25.4/2, 12.5, 'center');
    }
    if (txtCity.value == "León")
    {
      doc.text("Arq. Antonio Lavín Villa", 4.2333, 14.5);
      doc.text("Javier Everardo Hernández Moreno", 16.9333, 14.5);
      doc.setFontSize(8.5);
      doc.text("Representante Legal", 4.6333, 15);
      doc.text("Instructor del curso", 18.4333, 15);
    }
    else
    {
      doc.text("Arq. Antonio Lavín Villa", 25.4/2, 14.5, 'center');
      doc.setFontSize(8.5);
      doc.text("Representante Legal", 25.4/2, 15, 'center');
    }
    doc.setFontSize(12);
    doc.text("Otorga la presente constancia a:", 25.4/2, 5, 'center');
    doc.text("Por haber participado en el curso de:", 25.4/2, 7, 'center');
    doc.text(nombreComercial, 25.4/2, 9.5, 'center')
    var textWidth = doc.getStringUnitWidth(text) * doc.internal.getFontSize() / doc.internal.scaleFactor;
    var textOffset = (doc.internal.pageSize.getWidth() - (textWidth / 800)) / 2;
    const fontSize = 12;
    let startX = textOffset - 0.2;
    let startY = 9;
    const arrayOfNormalAndBoldText = txtInfo.split('**');
    arrayOfNormalAndBoldText.map((text, i) => {
      doc.setFont("calibri-bold");
    // every even item is a normal font weight item
    if (i % 2 === 0) {
      doc.setFont("calibri-normal");
      doc.text(text, startX, startY);
      startX = startX + ( (doc.getStringUnitWidth(text) / 800 ) / 2);
    }
    else
    {
      doc.text(text, startX, startY);
      startX = startX + ( (doc.getStringUnitWidth(text) / 2600 ) / 2);
    }
    });
    doc.setFontSize(22); // 22
    doc.setFont("calibri-normal");
    
    doc.text(sheetData[a][4] + " " + sheetData[a][2] + " " + sheetData[a][3], 25.4/2, 6, 'center');
    doc.text(sheetData[2][1], 25.4/2, 8, 'center');
    doc.setFontSize(10);
    doc.setTextColor( 206, 206, 206 )
    doc.text("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", 25.4/2, 15.8, 'center')
    //doc.fromHTML(htmlinfo.innerHTML, 5, 8.5)


    
  }
  doc.save(filename + '.pdf')
  console.log(sheetsLength);
  console.log(finished);
  if (finished >= sheetsLength - 1)
    {
      loader.style.display = "none"
    }
    else
    {
      finished++;
    }
  // addImage(image, format, X, Y, width, height)
  //htmlinfo.innerHTML = 'León, <b>Guanajuato</b>'
  //doc.fromHTML(htmlinfo, 15.5, 2)
}

var imageSelected = () => {
  var reader = new FileReader();
  reader.onloadend = function() {
    imgLogo.src = reader.result;
    imgLogo.style.height = "100px";
    imgLogo.style.width = "125px";
}
  reader.readAsDataURL(inputLogo.files[0]);
}
