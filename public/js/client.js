import { font_normal } from '../css/calibril-normal.js';
import { font_bold } from '../css/calibril-bold.js';
import { font } from '../css/opensans-bold.js';
var inputFile = document.getElementById("input");
var inputLogo = document.getElementById("inputLogo");
var imgLogo = document.getElementById("imgLogo");
var btnGenerateI = document.getElementById("btnGenerateI")
var btnGenerateG = document.getElementById("btnGenerateG")
var btnGenerateDC = document.getElementById("btnGenerateDC")
var firmaElectronica = document.getElementById("flexSwitchCheckDefault");
var switchVirtual = document.getElementById("switchVirtual");
var switchBomberos = document.getElementById("switchBomberos");
var imgLavi = new Image()
var firmaALV = new Image()
firmaALV.src = "./img/Firma ALV.PNG"
var firmaCapacitador = new Image()
firmaCapacitador.src = "./img/Firma Javier Capacitador.PNG"
imgLavi.src = "./img/LAVI Fire_HI-RES.jpg"
var imgEsquina = new Image()
imgEsquina.src = "./img/esquina/esquina-rojo.png"
var borde = new Image()
var displayError = document.getElementById("displayError")
var displaySuccess = document.getElementById("displaySuccess")
var loaderIcon = document.getElementById("loaderIcon")
loaderIcon.style.display = 'block'
displaySuccess.style.display = 'none';
displayError.style.display = 'none';


borde.src = "./img/borde/borde-rojo.png"
var loader = document.getElementById("loader");
var txtFile = document.getElementById("txtFile");
var codigos = [""];
var txtInfo = "";
var text = "";
var logoClient = new Image()
var image = new Image()
image.src = "./img/LAVI Fire Profile Pic.png"
var finished = 0;
var correo = new Image()
correo.src = "./img/Correo.png"
var telefono = new Image()
telefono.src = "./img/Telefono.png"
var ubicacion = new Image()
ubicacion.src = "./img/Ubicacion.png"
var file = null;

var alertSuccess = document.getElementById("alertSuccess");
alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente</span>"

inputFile.addEventListener("change", function() {
  file = inputFile.files[0];
  txtFile.innerHTML = file.name;
  txtFile.style.height = "auto"
  //else {
  //  alert("No se ha elegido un estado");
  //}
})

switchVirtual.addEventListener("change", () => {
  if(switchVirtual.checked) {
    if(switchBomberos.checked) {
      switchBomberos.checked = false;
    }
  }
})

switchBomberos.addEventListener("change", () => {
  if(switchBomberos.checked) {
    if(switchVirtual.checked) {
      switchVirtual.checked = false;
    }
  }
})

var generatePDF = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength, sheets) => {
  //console.log(firmaElectronica.checked)
  try {
  for (var a = 6; a < sheetData.length; a++)
  {
    if (a == 6)
    {
      if (sheetData[6][2] != null)
      {
        var doc = new jsPDF('l', 'pt', [540, 720]) //19.05, 25.4
        doc.addFileToVFS("calibril-normal.ttf", font_normal)
        doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
        doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
        doc.addFileToVFS("calibril-bold.ttf", font_bold)
        doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
        doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
        
      }
    }
    else
    {
      if (sheetData[a][2] != null)
      {
        doc.addPage(720, 540) //19.05, 25.4
      }
    }

    if(sheetData[a][2] != null)
    {
      doc.setFont("calibri-normal");
      doc.setTextColor( 0, 0, 0 )
      
      doc.addImage(imgEsquina, "PNG", 578.2677, 0, 141.732, 141.732); // 20.4, 0, 5, 5
      doc.addImage(borde, "PNG", 0, 496.063, 850.394, 42.5197); // 20.4, 0, 5, 5
      doc.addImage(imgLavi, "JPEG", 14.1732, 14.1732, 220.9448, 67.3622); // 0.5, 0.5, 9, 2.7
      doc.addImage(logoClient, "JPEG", 620, 452.5, 40, 40);
      doc.addImage(correo, "PNG", 310, 465, 15, 15);
      doc.addImage(telefono, "PNG", 435, 465, 15, 15);
      doc.addImage(ubicacion, "PNG", 170, 465, 15, 15);
      
      
      doc.addImage(image, "png", 56.6929, 452.5, 40, 40) // 2, 15.7, 1.8, 1.8
      doc.setFontSize(10);
      //doc.text(text, 720/2, 85.0394, 'center') // 25.4/2, 3
      doc.text(ciudadFecha, 651.969 - doc.getTextDimensions(ciudadFecha).w, 85.0394);
      doc.text("LAVI Fire Workshop México, S.A. de C.V.", 720/2, 322.2443, 'center'); // 25.4/2, 11
      doc.text("Capacitador en Materia de Protección Civil", 720/2, 336.4176, 'center'); // 25.4/2, 11.5
      doc.text(codigos[0], 720/2, 350.5903, 'center'); // 25.4/2, 12
      if (codigos.length > 1)
      {
        doc.text(codigos[1], 720/2, 364.764, 'center'); // 25.4/2, 12.5
      }
      if (sheetData[1][4] == "León")
      {
        doc.text("Arq. Antonio Lavín Villa", 119.999055, 424.0236); // 4.2333, 14.5
        doc.text("Javier Everardo Hernández Moreno", 479.9990551, 424.0236); // 16.9333, 14.5
        doc.setFontSize(8.5);
        doc.text("Representante Legal", 131.337638, 435.197); // 4.6333, 15
        doc.text("Instructor del curso", 522.5187402, 435.197); // 18.4333, 15
      }
      else
      {
        doc.text("Arq. Antonio Lavín Villa", 720/2, 424.0236, 'center'); // 25.4/2, 14.5
        doc.setFontSize(8.5);
        doc.text("Representante Legal", 720/2, 435.197, 'center'); // 25.4/2, 15
      }
      doc.setFontSize(12);
      doc.text("Otorga la presente constancia a:", 720/2, 121.831, 'center'); // 25.4/2, 5
      doc.text("Por haber participado en el curso de:", 720/2, 198.425, 'center'); // 25.4/2, 7
      doc.text(nombreComercial, 720/2, 294.291, 'center') // 25.4/2, 9.5
      var startX = 360 - (doc.getTextDimensions(text).w / 2);
      const arrayOfNormalAndBoldText = txtInfo.split('**');
      arrayOfNormalAndBoldText.map((mytext, i) => {
      // every even item is a normal font weight itemS
      doc.setFont("calibri-bold");
      // every even item is a normal font weight item
      if (i % 2 === 0) {
        doc.setFont("calibri-normal");
        doc.text(mytext, startX, 275.118);
        var dim = doc.getTextDimensions(mytext);
        if (i == 0) {
          startX = startX + dim.w + 2; 
        }
        else
        {
          //startX = startX + (dim.w / 7.2);
          startX = startX + dim.w + 5; 
        }
      }
      else
      {
        doc.text(mytext, startX, 275.118);
        var dim = doc.getTextDimensions(mytext);
        //startX = startX + (dim.w / 6.3);
        startX = startX + dim.w + 2; 
      }
      });
      startX = 0;
      doc.setFontSize(22); // 22
      doc.setFont("calibri-normal");
      
      doc.text(sheetData[a][8], 720/2, 160.128, 'center'); // 25.4/2, 6
      doc.text(sheetData[2][1], 720/2, 236.772, 'center'); // 25.4/2, 8
      doc.setFontSize(10);
      doc.setTextColor( 206, 206, 206 )
      doc.text("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", 720/2, 447.874, 'center') // 25.4/2, 15.8
      if (firmaElectronica.checked == true) {
        if (sheetData[1][4] == "León") {
          doc.addImage(firmaALV, "PNG", 89.999055, 344.0236, 175, 110)
          doc.addImage(firmaCapacitador, "PNG", 494.9990551, 349.0236, 120, 120)
        }
        else {
          doc.addImage(firmaALV, "PNG", 280, 344.0236, 175, 110)
        }
      }
      doc.setTextColor( 177, 177, 177 )
      doc.setFontSize(8.5);
      doc.text("lavifire@lavi.com.mx", 330, 475); // 360
      doc.text("O.    (477) 713 0016", 460, 465);
      doc.text("M1. (477) 670 2737", 460, 475);
      doc.text("M2. (477) 144 3124", 460, 485);
      doc.text("Av. Las Américas 401-B", 230, 465, 'center');
      doc.text("Col. Andrade CP 37020", 230, 475, 'center');
      doc.text("León, Guanajuato", 230, 485, 'center');
      //doc.fromHTML(htmlinfo.innerHTML, 5, 8.5) 
    }
  }
  if (sheetData[6][2] != null)
  {
    doc.save(filename + '.pdf')
  }
    console.log(sheetsLength)
  if (finished >= sheetsLength - 3)
    {
      loaderIcon.style.display = "none"
      displaySuccess.style.display = "block";
    }
    else
    {
      finished++;
    }
  }
  catch (error) {
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
  }
  // addImage(image, format, X, Y, width, height)
  //htmlinfo.innerHTML = 'León, <b>Guanajuato</b>'
  //doc.fromHTML(htmlinfo, 15.5, 2)
}

inputLogo.addEventListener("change", () => {
  var reader = new FileReader();
  reader.onloadend = function() {
    imgLogo.src = reader.result;
    logoClient.src = imgLogo.src
}
  reader.readAsDataURL(inputLogo.files[0]);
})


function wait(ms){
  var start = new Date().getTime();
  var end = start;
  while(end < start + ms) {
    end = new Date().getTime();
 }
}

btnGenerateI.addEventListener("click", () => 
{
  try {
    if (imgLogo.src.indexOf("img/404-error.png") == -1 && file != null){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(function(sheets) {
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Estados Municipios" && sheets[i - 1].name != "Direcciones" && sheets[i - 1].name != "RFCs") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              console.log(sheetData)
              if (sheetData[0][1] == null || sheetData[1][1] == null || sheetData[2][1] == null || sheetData[3][1] == null 
                && sheetData[4][1] == null || sheetData[0][4] == null || sheetData[1][4] == null || sheetData[2][4] == null
                && sheetData[3][4] == null || sheetData[2][6] == null || sheetData[6][2] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                return 
              }
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
              if (codigos.length > 1)
              {
                codigos.pop();
              }
              codigos[0] = sheetData[2][4];
              if (sheetData[3][4] != null)
              {
                codigos.push(sheetData[3][4]);
              }
  
              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var ciudadFecha = sheetData[1][4] + ", " + sheetData[0][4] + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - " + nombreCurso + " - " + nombreComercial
              // .getDate devuelve el dia
              // .getMonth devuelve el mes 0= Enero, 1 = Febrero, etc
              // .getFullYear devuelve el año completo ej: 1995
              generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets);
            })
          }
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
    else {
      document.getElementById("alertError").innerHTML='Falta logo y/o archivo excel por seleccionar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
  } catch(error) {
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
    document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
  }
  
}
)


btnGenerateDC.addEventListener("click", () => {
  try {
    if (imgLogo.src.indexOf("img/404-error.png") == -1 && file != null){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(function(sheets) {
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Estados Municipios" && sheets[i - 1].name != "Direcciones" && sheets[i - 1].name != "RFCs") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              console.log(sheetData)
              if (sheetData[0][1] == null || sheetData[1][1] == null || sheetData[2][1] == null || sheetData[3][1] == null 
                && sheetData[4][1] == null || sheetData[0][4] == null || sheetData[1][4] == null || sheetData[2][4] == null
                && sheetData[3][4] == null || sheetData[2][6] == null || sheetData[6][2] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                return 
              }
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
              if (codigos.length > 1)
              {
                codigos.pop();
              }
              codigos[0] = sheetData[2][4];
              if (sheetData[3][4] != null)
              {
                codigos.push(sheetData[3][4]);
              }
  
              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var razonSocial = sheetData[0]
              var date = fecha.getFullYear() + "-" + mesNumber + "-" + dia;
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - " + nombreCurso + " - " + nombreComercial
              generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date);
            })
          }
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
    else {
      document.getElementById("alertError").innerHTML='Falta logo y/o archivo excel por seleccionar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
  } catch(error) {
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
    document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
  }
})

btnGenerateG.addEventListener("click", () => {
  try {
    if (file != null){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(function(sheets) {
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Estados Municipios" && sheets[i - 1].name != "Direcciones" && sheets[i - 1].name != "RFCs") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              console.log(sheetData)
              if (sheetData[0][1] == null || sheetData[1][1] == null || sheetData[2][1] == null || sheetData[3][1] == null 
                && sheetData[4][1] == null || sheetData[0][4] == null || sheetData[1][4] == null || sheetData[2][4] == null
                && sheetData[3][4] == null || sheetData[2][6] == null || sheetData[6][2] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                return 
              }
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
              if (codigos.length > 1)
              {
                codigos.pop();
              }
              codigos[0] = sheetData[2][4];
              if (sheetData[3][4] != null)
              {
                codigos.push(sheetData[3][4]);
              }
  
              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var ciudadFecha = sheetData[1][4] + ", " + sheetData[0][4] + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
              var fechaConstancia = dia  + " de " + mes + " de " + fecha.getFullYear();
              var razonSocial = sheetData[0]
              //razonSocial[1]
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - " + nombreCurso + " - " + nombreComercial
              // .getDate devuelve el dia
              // .getMonth devuelve el mes 0= Enero, 1 = Febrero, etc
              // .getFullYear devuelve el año completo ej: 1995
              var participantes = [];
              var mult = false;
              var numC = 1;
              for (var a = 6; a < sheetData.length; a++) {
                if (sheetData[a][2] != null) {
                  if(participantes.length < 45) {
                    participantes.push(sheetData[a][8])
                  }
                  else {
                    mult = true;
                    if (sheetData[0][4] == "Jalisco") {
                      if(sheetData[2][1] == "Evacuación, Búsqueda y Rescate"){
                        nombreCurso = "Evacuación"
                        filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - Evacuación - " + nombreComercial
                        for (var i = 0; i < 2; i++) {
                          generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                          nombreCurso = "Búsqueda y Rescate"
                          filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - Búsqueda y Rescate - " + nombreComercial
                        }
                      }
                      else {
                        generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                      }
                    }
                    else {
                      generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                    }
                    participantes = [];
                    participantes.push(sheetData[a][8])
                    numC++;
                  }
                }
              }
              if (participantes.length > 0) {
                if (sheetData[0][4] == "Jalisco") {
                  if(sheetData[2][1] == "Evacuación, Búsqueda y Rescate"){
                    nombreCurso = "Evacuación"
                    filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - Evacuación - " + nombreComercial
                    for (var i = 0; i < 2; i++) {
                      generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                      nombreCurso = "Búsqueda y Rescate"
                      filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias - Búsqueda y Rescate - " + nombreComercial
                    }
                  }
                  else {
                    generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                  }
                }
                else {
                  generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantes, mult, numC);
                }
              }
            })
          }
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
    else {
      document.getElementById("alertError").innerHTML='Falta logo y/o archivo excel por seleccionar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
      document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
  } catch(error) {
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
    document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
  }
})

var generatePDFGrupal = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength, sheets, fechaConstancia, nombreCurso, participantes, mult, numC) => {
  try {
    if (sheetData[6][2] != null)
      {
        console.log("Entro")
        var doc = new jsPDF('p', 'pt', [612, 792]);
        doc.addFileToVFS("calibril-normal.ttf", font_normal)
        doc.addFileToVFS("open-sans-condensed.bold.ttf", font)
        doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
        doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
        doc.addFont('open-sans-condensed.bold.ttf', 'custom', 'normal');
        doc.addFileToVFS("calibril-bold.ttf", font_bold)
        doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
        doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
        doc.addImage(imgLavi, "JPEG", 57, 26, 149.9015, 42.6811); //220.9448 x 67.3622 
        if(firmaElectronica.checked == true) {
          doc.addImage(firmaALV, "PNG", 227, 564, 175, 110)
        }
        doc.setDrawColor(255, 0, 0);
        doc.setLineWidth(2.0);
        doc.line(57, 75, 557, 75); // donde empieza en X, x, donde termina en X, 87 527
        doc.setFont("calibri-normal");
        doc.setTextColor( 0, 0, 0 )
        doc.setFontSize(9.5);
        var dim = doc.getTextDimensions(ciudadFecha, {fontSize: 9.5});
        doc.text(ciudadFecha, 557 - dim.w, 105);
        doc.setFont("custom");
        doc.setFontSize(20.0);
        doc.text("CONSTANCIA DE CAPACITACION", 612/2, 140, 'center');
        doc.setFontSize(9.0);
        doc.setFont("calibri-normal");
        var info = "Hago constar que el personal que labora en " + razonSocial[1] + " (" + nombreComercial +"), ubicado en " + sheetData[4][1] + " "
        + sheetData[1][4] + ", " + sheetData[0][4] + ", participó de manera satisfactoria en el *CURSO BASICO DE " + nombreCurso.toUpperCase() 
        + ", *con una carga horaria de 08 h. "
        var info2 = "";
        if (switchVirtual.checked) {
          info2 = "La capacitación se impartió el día *"+ fechaConstancia +" *de manera virtual; con el programa y participantes siguientes: "
        }
        else if (switchBomberos.checked) {
          info2 = "La capacitación se impartió el día *"+ fechaConstancia +" *en las instalaciones del Campo de Entrenamiento para Bomberos, ubicado en Prado de los Nardos No. 1016. Col. San Gaspar de las Flores, Tonalá, Jalisco; con el programa y participantes siguientes: "
        }
        else {
          info2 = "La capacitación se impartió el día *"+ fechaConstancia +" *en las instalaciones de " + nombreComercial + "; con el programa y participantes siguientes: "
        }
        
        const arrayOfNormalAndBoldText = info.split('');
        const arrayOfNormalAndBoldText2 = info2.split('');

        var startX = 90; // 120
        var BoldFont = false;
        var startY = 160;
        var palabra = "";
        var letrasMayusculas = 0;
        var invalid = false;
        arrayOfNormalAndBoldText.map((mytext, i) => {
          if (mytext == ' ' || mytext == '*') {
            if (mytext == ' ') {
              if (doc.getTextDimensions(palabra, {fontSize: 9.0}).w + startX < 557) { // 527
                doc.text(palabra, startX, startY);
                //if(palabra[palabra.length -1])
                if (palabra == palabra.toUpperCase()) {
                  for (var b = 0; b < palabra.length; b++) {
                    if (palabra[b] >= '0' && palabra[b] <= '9') {
                      invalid = true;
                      letrasMayusculas = 0;
                    }
                    else {
                      if (invalid == false) {
                        letrasMayusculas++;
                      }
                    }
                  }
                }
                startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2 - (letrasMayusculas * 0.8);
                palabra = "";
                letrasMayusculas = 0;
                palabra = "";
                invalid = false;
              }
              else {
                startX = 57; // 95
                startY += 11;
                doc.text(palabra, startX, startY);
                if (palabra == palabra.toUpperCase()) {
                  for (var b = 0; b < palabra.length; b++) {
                    if (palabra[b] >= '0' && palabra[b] <= '9') {
                      invalid = true;
                      letrasMayusculas = 0;
                    }
                    else {
                      if (invalid == false) {
                        letrasMayusculas++;
                      }
                    }
                  }
                }
                startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2 - (letrasMayusculas * 0.6);
                palabra = "";
                letrasMayusculas = 0;
                invalid = false;
              }
            }
            else {
              BoldFont = !BoldFont;
              if (BoldFont) {
                doc.setFont("calibri-bold");
      
              }
              else {
                doc.setFont("calibri-normal");
              }
            }
          }
          else {
            palabra += mytext;
          }
        });

        startX = 90; // 120
        startY = 200;
        arrayOfNormalAndBoldText2.map((mytext, i) => {
          if (mytext == ' ' || mytext == '*') {
            if (mytext == ' ') {
              if (doc.getTextDimensions(palabra, {fontSize: 9.0}).w + startX < 557) { // 527
                doc.text(palabra, startX, startY);
                //if(palabra[palabra.length -1])
                startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2;
                if (palabra.indexOf(";") != -1) {
                  startX -= 1;
                }
                palabra = "";
              }
              else {
                startX = 57; // 95
                startY += 11;
                doc.text(palabra, startX, startY);
                startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2;
                if (palabra.indexOf(";") != -1) {
                  startX -= 1;
                }
                palabra = "";
              }
            }
            else {
              BoldFont = !BoldFont;
              if (BoldFont) {
                doc.setFont("calibri-bold");
            
              }
              else {
                doc.setFont("calibri-normal");
              }
            }
          }
          else {
            palabra += mytext;
          }
        });

        doc.setFont("calibri-bold");
    
        doc.text("PROGRAMA: CURSO BASICO DE " + nombreCurso.toUpperCase(), 57, 235) // 120
        doc.setFont("calibri-normal");
        switch(nombreCurso) {
          case "Búsqueda y Rescate":
            doc.text("a.", 90, 246)
            doc.text("Introducción. ¿Qué es Búsqueda y Rescate?", 107, 246)
            doc.text("b.", 90, 257)
            doc.text("Búsqueda y Rescate Urbano.", 107, 257)
            doc.text("a.", 120, 268)
            doc.text("Búsqueda y Rescate en estructuras colapsadas.", 137, 268)
            doc.text("c.", 90, 279)
            doc.text("Organización e inicio de una operación.", 107, 279)
            doc.text("a.", 120, 290)
            doc.text("Etapas para la respuesta de una operación.", 137, 290)
            doc.text("d.", 90, 301)
            doc.text("Búsqueda y Localización.", 107, 301)
            doc.text("e.", 90, 312)
            doc.text("Consideraciones de Seguridad.", 107, 312)
            doc.text("a.", 120, 323)
            doc.text("Equipo de seguridad personal.", 137, 323)
            break;
          case "Primeros Auxilios":
            doc.text("a.", 90, 246)
            doc.text("Principios Generales de los Primeros Auxilios.", 107, 246)
            doc.text("b.", 90, 257)
            doc.text("Evaluación Primaria.", 107, 257)
            doc.text("c.", 90, 268)
            doc.text("Soporte Básico de Vida.", 107, 268)
            doc.text("d.", 90, 279)
            doc.text("Atención General a Hemorragias.", 107, 279)
            doc.text("e.", 90, 290)
            doc.text("Manejo del Estado de Shock.", 107, 290)
            doc.text("f.", 90, 301)
            doc.text("Heridas y Quemaduras.", 107, 301)
            doc.text("g.", 90, 312)
            doc.text("Lesiones Traumáticas en Huesos.", 107, 312)
            doc.text("h.", 90, 323)
            doc.text("Movilización y traslado de lesionados.", 107, 323)
            break;
          case "Prevención y Combate de Incendios": 
            doc.text("a.", 90, 246)
            doc.text("Conceptos Generales para la Prevención y Combate de Incendios.", 107, 246)
            doc.text("b.", 90, 257)
            doc.text("Química del Fuego.", 107, 257)
            doc.text("c.", 90, 268)
            doc.text("Prevención y Control de Incendio.", 107, 268)
            doc.text("a.", 120, 279)
            doc.text("Clasificación del fuego.", 137, 279)
            doc.text("b.", 120, 290)
            doc.text("Agentes extinguidores del fuego.", 137, 290)
            doc.text("d.", 90, 301)
            doc.text("Acciones durante un incendio.", 107, 301)
            doc.text("a.", 120, 312)
            doc.text("Uso y manejo de extintores.", 137, 312)
            doc.text("b.", 120, 323)
            doc.text("¿Qué hacer en caso de un incendio?", 137, 323)
            break;
          case "Evacuación":
            doc.text("a.", 90, 246)
            doc.text("Generalidades.", 107, 246)
            doc.text("a.", 120, 257)
            doc.text("Agentes perturbadores.", 137, 257)
            doc.text("b.", 120, 268)
            doc.text("Principios básicos de acción.", 137, 268)
            doc.text("b.", 90, 279)
            doc.text("Integración y actividades. Antes, Durante y Después de la Emergencia.", 107, 279)
            doc.text("a.", 120, 290)
            doc.text("Identificación del personal brigadista.", 137, 290)
            doc.text("b.", 120, 301)
            doc.text("Equipo de seguridad personal.", 137, 301)
            doc.text("c.", 90, 312)
            doc.text("Simulacros.", 107, 312)
            doc.text("d.", 90, 323)
            doc.text("Caso Práctico.", 107, 323)
            break;
          default:
            doc.text("a.", 90, 246)
            doc.text("Integración y actividades. Antes, Durante y Después de la Emergencia.", 107, 246)
            doc.text("a.", 120, 257)
            doc.text("Identificación del personal brigadista.", 137, 257)
            doc.text("b.", 120, 268)
            doc.text("Equipo de seguridad personal.", 137, 268)
            doc.text("b.", 90, 279)
            doc.text("Simulacros.", 107, 279)
            doc.text("c.", 90, 290)
            doc.text("Caso Práctico Simulacro de Evacuación / Repliegue.", 107, 290)
            doc.text("d.", 90, 301)
            doc.text("¿Qué es Búsqueda y Rescate?", 107, 301)
            doc.text("e.", 90, 312)
            doc.text("Búsqueda y Rescate Urbano.", 107, 312)
            doc.text("f.", 90, 323)
            doc.text("Búsqueda y Localización.", 107, 323)
            break;
        }
        doc.setFont("calibri-bold");
    
        doc.text("IMPARTIDO POR EL INSTRUCTOR:", 57, 345)

        doc.setFont("calibri-normal");
        doc.text("T.B.G.I.R. Arq. Antonio Lavín Villa. (LAVI Fire Workshop México, S.A. de C.V.)", 90, 356)

        doc.setFont("calibri-bold");
    
        doc.text("PERSONAL APROBADO:", 57, 375)
        doc.setFillColor(0, 0, 0);
        startX = 60;
        startY = 385; 
        doc.setFont("calibri-normal");
        for (var a = 0; a < participantes.length; a++) {
          if (participantes[a] != null) {
            doc.circle(startX, startY, 2, 'F');
            doc.text(participantes[a], startX + 15, startY +3)
            if (startY + 10 <= 525) {
              startY += 10;
            }
            else {
              startY = 385;
              startX += 180;
            }
          }
        }
        doc.setFont("calibri-bold");
        doc.text("Importante:", 57, 550)
        doc.setFont("calibri-normal");
        doc.text("El presente documento tiene 1 año de validez a partir de la fecha de expedición.", 107, 550)
        doc.text("Atentamente", 612/2, 570, 'center')
        doc.text("Arq. Antonio Lavín Villa", 612/2, 640, 'center')
        doc.text("LAVI Fire Workshop México, S.A. de C.V.", 612/2, 650, 'center')
        doc.text("Capacitador en Materia de Protección Civil", 612/2, 660, 'center')
        doc.text(sheetData[2][4], 612/2, 670, 'center')
        doc.text("REG. DE S.T.P.S LFW-160516-NJ1-0013", 612/2, 680, 'center')
        doc.setFontSize(7.5);
        doc.text("Se emite la presente, bajo lo dispuesto por los Artículos 19 Fracción XVII, 41, 42, 43 Fracción I, II, III, IV, V, VI, 45,56, de la Ley General de Protección Civil,", 612/2, 695, 'center')
        doc.text("98 del Reglamento de la Ley General de Protección Civil.", 612/2, 705, 'center')
        doc.setFontSize(9.0);
        doc.setDrawColor(255, 0, 0);
        doc.line(57, 715, 557, 715);
        doc.text("Av. Las Américas 401-B | Col. Andrade CP 37020 | León, Guanajuato",  612/2, 726, 'center')
        doc.setFont("calibri-bold");
        doc.text("O.",  194, 737)
        doc.setFont("calibri-normal");
        doc.text("(477) 713 0016 |",  204, 737)
        doc.setFont("calibri-bold");
        doc.text("M1.",  267, 737)
        doc.setFont("calibri-normal");
        doc.text("(477) 670 2737 |",  284, 737)
        doc.setFont("calibri-bold");
        doc.text("M2.",  347, 737)
        doc.setFont("calibri-normal");
        doc.text("(477) 144 3124",  364, 737)
        doc.text("lavifire@lavi.com.mx",  612/2, 748, 'center')
      }
    
    if (sheetData[6][2] != null)
    {
      if (mult == true) {
        doc.save(filename + '_' + numC  + '.pdf')
      }
      else {
        doc.save(filename +'.pdf')
      }
    }
      console.log(finished)
    if (finished >= sheetsLength - 3)
      {
        loaderIcon.style.display = "none"
        displaySuccess.style.display = "block";
      }
      else
      {
        finished++;
      }
    }
    catch (error) {
      document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
    }
}

var generatePDFDC3 = (razonSocial, sheetData, filename, sheetsLength, sheets, date) => {
  try {
    for (var a = 6; a < sheetData.length; a++)
    {
      if (a == 6)
      {
        if (sheetData[6][2] != null)
        {
          var doc = new jsPDF('p', 'pt', [612, 792]);
          doc.addFileToVFS("calibril-normal.ttf", font_normal)
          doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
          doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
          doc.addFileToVFS("calibril-bold.ttf", font_bold)
          doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
          doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
          
        }
      }
      else
      {
        if (sheetData[a][2] != null)
        {
          doc.addPage(720, 540) //19.05, 25.4
        }
      }
  
      if(sheetData[a][2] != null)
      {
        doc.setFont("calibri-normal");
        doc.setTextColor( 0, 0, 0 )
        
        doc.addImage(imgLavi, "JPEG", 14.1732, 14.1732, 220.9448, 67.3622); // 0.5, 0.5, 9, 2.7
        doc.addImage(logoClient, "JPEG", 620, 452.5, 40, 40); 
      }
    }
    if (sheetData[6][2] != null)
    {
      doc.save(filename + '.pdf')
    }
      console.log(sheetsLength)
    if (finished >= sheetsLength - 3)
      {
        loaderIcon.style.display = "none"
        displaySuccess.style.display = "block";
      }
      else
      {
        finished++;
      }
    }
    catch (error) {
      document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
    }
}