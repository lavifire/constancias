import { font_normal } from '../css/calibril-normal.js';
import { font_bold } from '../css/calibril-bold.js';
import { font } from '../css/opensans-bold.js';
var doc = null;
var inputFile = document.getElementById("input");
var inputLogo = document.getElementById("inputLogo");
var imgLogo = document.getElementById("imgLogo");
var btnGenerateI = document.getElementById("btnGenerateI")
var btnGenerateG = document.getElementById("btnGenerateG")
var btnGenerateDC = document.getElementById("btnGenerateDC")
var src;
var vencido = false;
var firmaElectronica = document.getElementById("flexSwitchCheckDefault");
var switchVirtual = document.getElementById("switchVirtual");
var switchBomberos = document.getElementById("switchBomberos");
var cursoName = ";"
var registrosData;
var imgLavi = new Image()
var firmaALV = new Image()
var firmaREP1 = new Image()
var firmaREP2 = new Image()
firmaALV.src = "./img/Firma ALV.jpg"
firmaREP1.src = "./img/Firma Representantes Suburbia1.jpg"
firmaREP2.src = "./img/Firma Representantes Suburbia2.jpg"
var firmaCapacitador = new Image()
firmaCapacitador.src = "./img/Firma Javier Capacitador.jpg"
imgLavi.src = "./img/LAVI Fire_HI-RES.jpg"
var imgEsquina = new Image()
imgEsquina.src = "./img/esquina/esquina-rojo.jpg"
var borde = new Image()
var displayError = document.getElementById("displayError")
var displaySuccess = document.getElementById("displaySuccess")
var loaderIcon = document.getElementById("loaderIcon")
loaderIcon.style.display = 'block'
displaySuccess.style.display = 'none';
displayError.style.display = 'none';
var duracion = [];


borde.src = "./img/borde/borde-rojo.jpg"
var loader = document.getElementById("loader");
var txtFile = document.getElementById("txtFile");
var codigos = [""];
var txtInfo = "";
var text = "";
var logoClient = new Image()
var image = new Image()
image.src = "./img/LAVI Fire Profile Pic.jpg"
var finished = 0;
var correo = new Image()
correo.src = "./img/Correo.jpg"
var telefono = new Image()
telefono.src = "./img/Telefono.jpg"
var ubicacion = new Image()
ubicacion.src = "./img/Ubicacion.jpg"
var file = null;
var path;
var ErrorFounded = false;

var alertSuccess = document.getElementById("alertSuccess");
alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente</span>"

inputFile.addEventListener("change", async function(event) {
  console.log(event.target.value)
  ErrorFounded = false;
  file = inputFile.files[0];
  path = URL.createObjectURL(file)
  console.log(path)
  console.log(file)
  var sheets = null;
  await readXlsxFile(file, { getSheets: true }).then(async function(mysheets) {
    sheets = mysheets;
  })
  
  for (var i = 0; i < sheets.length; i++) {
    await readXlsxFile(file, { sheet: i + 1 }).then(async function(sheetData) {
      if (sheets[i].name == "Cursos PC") {
        for(var a = 3; a < sheetData.length; a++){
          duracion.push({"curso": sheetData[a][2], "duracion": sheetData[a][5]});
          //console.log(duracion[a - 3].curso);
        }
      }
      else if (sheets[i].name == "Registros PC") {
        await readXlsxFile(file, { sheet: i + 1 }).then(function(data) {
          registrosData = data;
        })
      }
      else if (sheets[i].name == "Evacuación, Búsqueda y Rescate" && sheetData[1][4] == "Jalisco") {
        await readXlsxFile(file, { sheet: i + 1 }).then(function(data) {
          for (var a = 7; sheetData[a][1] != null ; a++) {
            if (sheetData[a][11] == "" || sheetData[a][11] == null) {
              ErrorFounded = true;
            }
          }
        })
      }
    })
  }
  await new Promise(r => setTimeout(r, 1000));
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

var generatePDF = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength, sheets, cursoNombre) => {
  //console.log(firmaElectronica.checked)
  try {
    var JaliscoCondition = false;
    doc = null;
  for (var a = 7; a < sheetData.length; a++)
  {
    JaliscoCondition = false;

    if (cursoNombre =="Evacuación de Inmuebles"){
      if(sheetData[a][11] == "CBYR"){
        JaliscoCondition = true;
      }
    }
    else if (cursoNombre =="Busqueda y Rescate"){
      if(sheetData[a][11] != "CBYR"){
        JaliscoCondition = true;
      }
    }

    if(sheetData[a][1] != null && JaliscoCondition == false)
    {
      if (doc != null){
        doc.addPage(720, 540) //19.05, 25.4
      }
      else {
        doc = new jsPDF('l', 'pt', [540, 720]) //19.05, 25.4
        doc.addFileToVFS("calibril-normal.ttf", font_normal)
        doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
        doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
        doc.addFileToVFS("calibril-bold.ttf", font_bold)
        doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
        doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
      }
      doc.setFont("calibri-normal");
      doc.setTextColor( 0, 0, 0 )
      
      doc.addImage(imgEsquina, "JPEG", 578.2677, 0, 141.732, 141.732); // 20.4, 0, 5, 5
      doc.addImage(borde, "JPEG", 0, 496.063, 850.394, 42.5197); // 20.4, 0, 5, 5
      doc.addImage(imgLavi, "JPEG", 14.1732, 14.1732, 220.9448, 67.3622); // 0.5, 0.5, 9, 2.7
      doc.addImage(logoClient, "JPEG", 620, 452.5, 40, 40);
      doc.addImage(correo, "JPEG", 310, 465, 15, 15);
      doc.addImage(telefono, "JPEG", 435, 465, 15, 15);
      doc.addImage(ubicacion, "JPEG", 170, 465, 15, 15);
      
      
      doc.addImage(image, "JPEG", 56.6929, 452.5, 40, 40) // 2, 15.7, 1.8, 1.8
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
      doc.text(cursoName, 720/2, 236.772, 'center'); // 25.4/2, 8
      doc.setFontSize(10);
      if (firmaElectronica.checked == true) {
        if (sheetData[2][4] == "León") {
          doc.addImage(firmaALV, "JPEG", 89.999055, 344.0236, 175, 110)
          doc.addImage(firmaCapacitador, "JPEG", 494.9990551, 349.0236, 120, 120)
        }
        else {
          doc.addImage(firmaALV, "JPEG", 280, 344.0236, 175, 110)
        }
        doc.text(codigos[0], 720/2, 350.5903, 'center'); // 25.4/2, 12
        if (codigos.length > 1)
        {
          doc.text(codigos[1], 720/2, 364.764, 'center'); // 25.4/2, 12.5
        }
      }
      if (sheetData[2][4] == "León")
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
      doc.setTextColor( 206, 206, 206 )
      doc.text("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------", 720/2, 447.874, 'center') // 25.4/2, 15.8
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
  if (sheetData[7][1] != null)
  {
    doc.save(filename + '.pdf')
  }
    console.log(sheetsLength)
    console.log(finished)
  if (finished >= sheetsLength - 2)
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
    console.log(error);
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
  }
  // addImage(image, format, X, Y, width, height)
  //htmlinfo.innerHTML = 'León, <b>Guanajuato</b>'
  //doc.fromHTML(htmlinfo, 15.5, 2)
}

inputLogo.addEventListener("change", async () => {
  src = URL.createObjectURL(inputLogo.files[0])
  var imageTemp = new Image();
  imageTemp.src = src;
  await new Promise(r => setTimeout(r, 1000));
  if (imageTemp.width < 3522) {
    imgLogo.style.width = "230px"
    imgLogo.style.height = "200px"
  }
  else {
    imgLogo.style.width = "652px"
    imgLogo.style.height = "193px"
  }
  imgLogo.src = src
  logoClient.src = src
})

btnGenerateI.addEventListener("click",() => 
{
  try {
    vencido = false;
    displaySuccess.style.backgroundColor = "#04AA6D";
    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente</span>"
    if (imgLogo.src.indexOf("img/404-error.png") == -1 && file != null && imgLogo.style.height == "200px"){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(async function(sheets) {
        if (sheets.length < 6 || ErrorFounded == true) {
          if (ErrorFounded == true) {
            document.getElementById("alertError").innerHTML='Faltan datos en el excel'
          }
          else {
            document.getElementById("alertError").innerHTML='Formato de listado incorrecto'
          }
          loaderIcon.style.display = "none"
          loader.style.display = "block"
          displayError.style.display = "block"
          document.getElementById('imgLogo').src = './img/404-error.png'; 
          document.getElementById('input').files = null; 
          document.getElementById('txtFile').innerHTML = ''; 
          document.getElementById('txtFile').style.height = '25px';
          return;
        }
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Cursos PC" && sheets[i - 1].name != "Detalle" && sheets[i - 1].name != "ListaD") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              if (sheetData[2][1] == null || sheetData[4][1] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                return 
              }
              var date = sheetData[4];
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
              codigos[0] = sheetData[3][4];
              if (sheetData[4][4] != null)
              {
                codigos.push(sheetData[4][4]);
              }

              for (var a = 3; a < registrosData.length; a++) {
                if (codigos[0] == registrosData[a][4] || codigos[1] == registrosData[a][4]) {
                  if (registrosData[a][6] == "Vencido") {
                    displaySuccess.style.backgroundColor = "#af9003"
                    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente pero con los registros vencido</span>"
                    vencido = true;
                  }
                }
              }

              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var ciudadFecha = sheetData[2][4] + ", " + sheetData[1][4] + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
              var razonSocial = sheetData[1]
              txtInfo = 'Impartido el día **'+ dia + " de " + mes + " de " + fecha.getFullYear() +'**, para personal de **'+ razonSocial[1]
              text = 'Impartido el día '+ dia + " de " + mes + " de " + fecha.getFullYear() +', para personal de '+ razonSocial[1]
              var nombreComercial = sheetData[2][1]
              var nombreCurso = sheetData[3][1]
              cursoName = nombreCurso;
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias Individuales - " + nombreCurso + " - " + nombreComercial
              // .getDate devuelve el dia
              // .getMonth devuelve el mes 0= Enero, 1 = Febrero, etc
              // .getFullYear devuelve el año completo ej: 1995
              if (sheetData[3][1] == "Evacuación, Búsqueda y Rescate" && sheetData[1][4] == "Jalisco") {
                nombreCurso = "Evacuación de Inmuebles";
                cursoName = nombreCurso;
                filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias Individuales - " + nombreCurso + " - " + nombreComercial
                generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, nombreCurso);
                nombreCurso = "Busqueda y Rescate";
                cursoName = nombreCurso;
                var anio = fecha.getFullYear();
                dia = (fecha.getDate() >= 10) ? (fecha.getDate() + 1) : ("0" + (fecha.getDate() + 1))
                switch(mesNumber) {
                  case "01":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Febrero"
                      mesNumber = "02"
                    }
                    break;
                    case "02":
                      var bisiesto = false;
                      var aniosRestantes = fecha.getFullYear();
                      do {
                        if (aniosRestantes == 2000) {
                          bisiesto = true;
                        }
                        aniosRestantes = aniosRestantes - 4;
                      }
                      while (aniosRestantes >= 2000)
                      if (bisiesto == true) {
                        if (dia > 29) {
                          dia = "01";
                          mes = "Marzo"
                          mesNumber = "03"
                        }
                      }
                      else {
                        if (dia > 28) {
                          dia = "01";
                          mes = "Marzo"
                          mesNumber = "03"
                        }
                      }
                    break;
                    case "03":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Abril"
                      mesNumber = "04"
                    }
                    break;
                    case "04":
                    if (dia > 30) {
                      dia = "01";
                      mes = "Mayo"
                      mesNumber = "05"
                    }
                    break;
                    case "05":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Junio"
                      mesNumber = "06"
                    }
                    break;
                    case "06":
                    if (dia > 30) {
                      dia = "01";
                      mes = "Julio"
                      mesNumber = "07"
                    }
                    break;
                    case "07":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Agosto"
                      mesNumber = "08"
                    }
                    break;
                    case "08":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Septiembre"
                      mesNumber = "09"
                    }
                    break;
                    case "09":
                    if (dia > 30) {
                      dia = "01";
                      mes = "Octubre"
                      mesNumber = "10"
                    }
                    break;
                    case "10":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Noviembre"
                      mesNumber = "11"
                    }
                    break;
                    case "11":
                    if (dia > 30) {
                      dia = "01";
                      mes = "Diciembre"
                      mesNumber = "12"
                    }
                    break;
                    case "12":
                    if (dia > 31) {
                      dia = "01";
                      mes = "Enero"
                      mesNumber = "01"
                      anio = fecha.getFullYear() + 1;
                    }
                    break;
                }
                var ciudadFecha = sheetData[2][4] + ", " + sheetData[1][4] + " a " + dia  + " de " + mes + " de " + anio;
                txtInfo = 'Impartido el día **'+ dia + " de " + mes + " de " + anio +'**, para personal de **'+ razonSocial[1]
                text = 'Impartido el día '+ dia + " de " + mes + " de " + anio +', para personal de '+ razonSocial[1]
                filename = anio.toString().substr(-2) + mesNumber + dia + "_Constancias Individuales - " + nombreCurso + " - " + nombreComercial
                generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, nombreCurso);
              }
              else {
                generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, sheetData[3][1]);
              }
            })
          }
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png';
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px";  
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
    else {
      document.getElementById("alertError").innerHTML='Logo erroneo (faltante o formato incorrecto) y/o archivo excel por seleccionar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
      document.getElementById('imgLogo').src = './img/404-error.png'; 
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
  } catch(error) {
    console.log(error)
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
    vencido = false;
    displaySuccess.style.backgroundColor = "#04AA6D";
    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente</span>"
    if (imgLogo.src.indexOf("img/404-error.png") == -1 && file != null && imgLogo.style.height == "193px"){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(function(sheets) {
        if (sheets.length < 6 || ErrorFounded == true) {
          if (ErrorFounded == true) {
            document.getElementById("alertError").innerHTML='Faltan datos en el excel'
          }
          else {
            document.getElementById("alertError").innerHTML='Formato de listado incorrecto'
          }
          loaderIcon.style.display = "none"
          loader.style.display = "block"
          displayError.style.display = "block"
          document.getElementById('imgLogo').src = './img/404-error.png'; 
          document.getElementById('input').files = null; 
          document.getElementById('txtFile').innerHTML = ''; 
          document.getElementById('txtFile').style.height = '25px';
          ErrorFounded = false;
          return;
        }
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Cursos PC" && sheets[i - 1].name != "Detalle" && sheets[i - 1].name != "ListaD") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              console.log(sheetData)
              if (sheetData[2][1] == null || sheetData[4][1] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                
                return 
              }
              var date = sheetData[4];
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
              codigos[0] = sheetData[3][4];
              if (sheetData[4][4] != null)
              {
                codigos.push(sheetData[4][4]);
              }

              for (var a = 3; a < registrosData.length; a++) {
                if (codigos[0] == registrosData[a][4] || codigos[1] == registrosData[a][4]) {
                  if (registrosData[a][6] == "Vencido") {
                    displaySuccess.style.backgroundColor = "#af9003"
                    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente pero con los registros vencido</span>"
                    vencido = true;
                  }
                }
              }
  
              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var razonSocial = sheetData[1]
              var date = fecha.getFullYear() + "-" + mesNumber + "-" + dia;
              var nombreComercial = sheetData[2][1]
              var nombreCurso = sheetData[3][1]
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - " + nombreCurso + " - " + nombreComercial
              if (sheetData[3][1] == "Evacuación, Búsqueda y Rescate"){
                if(sheetData[1][4] == "Jalisco") {
                  console.log("Entro a condicion especial")
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Evacuación de Inmuebles - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Evacuación de Inmuebles", true);
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Busqueda y Rescate - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Busqueda y Rescate", true);
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Comunicación - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Comunicación", false);
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Formación de Brigadas e Introducción a la Protección Civil - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Formación de Brigadas e Introducción a la Protección Civil", false);
                }
                else {
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, sheetData[3][1], false);
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Comunicación - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Comunicación", false);
                  var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias DC3 - Formación de Brigadas e Introducción a la Protección Civil - " + nombreComercial
                  generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, "Formación de Brigadas e Introducción a la Protección Civil", false);
                }
              }
              else {
                generatePDFDC3(razonSocial, sheetData, filename, sheets.length - 2, sheets, date, sheetData[3][1], false);
              }
            })
          }
          //else {
          //  if (sheets[i - 1].name == "Registros PC") {
          //    readXlsxFile(file, { sheet: i }).then(function(data) {
          //      registrosData = data;
          //    })
          //  }
          //}
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png'; 
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    }
    else {
      document.getElementById("alertError").innerHTML='Logo erroneo (faltante o formato incorrecto) y/o archivo excel por seleccionar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
      document.getElementById('imgLogo').src = './img/404-error.png'; 
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    duracion = [];
    }
  } catch(error) {
    document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
    loaderIcon.style.display = "none"
    loader.style.display = "block"
    displayError.style.display = "block"
    document.getElementById('imgLogo').src = './img/404-error.png'; 
    document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
    duracion = [];
  }
})

btnGenerateG.addEventListener("click", () => {
  try {
    vencido = false;
    displaySuccess.style.backgroundColor = "#04AA6D";
    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente</span>"
    if (file != null){
      loader.style.display = "block"
      loaderIcon.style.display = "block"
      finished = 0;
      readXlsxFile(file, { getSheets: true }).then(function(sheets) {
        if (sheets.length < 6 || ErrorFounded == true) {
          if (ErrorFounded == true) {
            document.getElementById("alertError").innerHTML='Faltan datos en el excel'
          }
          else {
            document.getElementById("alertError").innerHTML='Formato de listado incorrecto'
          }
          loaderIcon.style.display = "none"
          loader.style.display = "block"
          displayError.style.display = "block"
          document.getElementById('imgLogo').src = './img/404-error.png'; 
          document.getElementById('input').files = null; 
          document.getElementById('txtFile').innerHTML = ''; 
          document.getElementById('txtFile').style.height = '25px';
          return;
        }
        for (var i = sheets.length; i > 0; i--) {
          if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Cursos PC" && sheets[i - 1].name != "Detalle" && sheets[i - 1].name != "ListaD") {
            readXlsxFile(file, { sheet: i }).then(function(sheetData) {
              console.log(sheetData)
              if (sheetData[2][1] == null || sheetData[4][1] == null){
                document.getElementById("alertError").innerHTML='Faltan datos en el excel'
                loaderIcon.style.display = "none";
                displayError.style.display = "block";
                return 
              }
              var date = sheetData[4];
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
              codigos[0] = sheetData[3][4];
              if (sheetData[4][4] != null)
              {
                codigos.push(sheetData[4][4]);
              }

              for (var a = 3; a < registrosData.length; a++) {
                if (codigos[0] == registrosData[a][4] || codigos[1] == registrosData[a][4]) {
                  if (registrosData[a][6] == "Vencido") {
                    displaySuccess.style.backgroundColor = "#af9003"
                    alertSuccess.innerHTML = "<strong>Success!</strong><span>Se crearon las constancias correctamente pero con los registros vencido</span>"
                    vencido = true;
                  }
                }
              }
  
              var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
              var ciudadFecha = sheetData[2][4] + ", " + sheetData[1][4] + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
              var fechaConstancia = dia  + " de " + mes + " de " + fecha.getFullYear();
              var razonSocial = sheetData[1]
              //razonSocial[1]
              var nombreComercial = sheetData[2][1]
              var nombreCurso = sheetData[3][1]
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
              var filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias Grupales - " + nombreCurso + " - " + nombreComercial
              // .getDate devuelve el dia
              // .getMonth devuelve el mes 0= Enero, 1 = Febrero, etc
              // .getFullYear devuelve el año completo ej: 1995
              var participantes = [];
              var participantesEI = [];
              var participantesBR = [];
              var mult = false;
              var numC = 1;
              for (var a = 7; a < sheetData.length; a++) {
                if (sheetData[a][1] != null) {
                  if(participantes.length < 45) {
                    participantes.push(sheetData[a][8])
                    if (sheetData[a][11] != null) {
                      if (sheetData[a][11].includes("CBYR") == true) {
                        participantesBR.push(sheetData[a][8])
                      }
                      else {
                        participantesEI.push(sheetData[a][8])
                      }
                    }
                    else {
                      // participantesEI.push(sheetData[a][8])
                    }
                  }
                  else {
                    mult = true;
                    if (sheetData[1][4] == "Jalisco") {
                      if(sheetData[3][1] == "Evacuación, Búsqueda y Rescate"){
                        nombreCurso = "Evacuación de Inmuebles"
                        filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias Grupales - Evacuación - " + nombreComercial
                        for (var i = 0; i < 2; i++) {
                          if(nombreCurso == "Evacuación de Inmuebles"){
                            generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantesEI, mult, numC);
                          }
                          else {
                            generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantesBR, mult, numC);
                          }
                          nombreCurso = "Búsqueda y Rescate"
                          var anio = fecha.getFullYear();
                          dia = (fecha.getDate() >= 10) ? (fecha.getDate() + 1) : ("0" + (fecha.getDate() + 1))
                          switch(mesNumber) {
                            case "01":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Febrero"
                              }
                              break;
                              case "02":
                                var bisiesto = false;
                                var aniosRestantes = fecha.getFullYear();
                                do {
                                  if (aniosRestantes == 2000) {
                                    bisiesto = true;
                                  }
                                  aniosRestantes = aniosRestantes - 4;
                                }
                                while (aniosRestantes >= 2000)
                                if (bisiesto == true) {
                                  if (dia > 29) {
                                    dia = "01";
                                    mes = "Marzo"
                                  }
                                }
                                else {
                                  if (dia > 28) {
                                    dia = "01";
                                    mes = "Marzo"
                                  }
                                }
                              break;
                              case "03":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Abril"
                              }
                              break;
                              case "04":
                              if (dia > 30) {
                                dia = "01";
                                mes = "Mayo"
                              }
                              break;
                              case "05":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Junio"
                              }
                              break;
                              case "06":
                              if (dia > 30) {
                                dia = "01";
                                mes = "Julio"
                              }
                              break;
                              case "07":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Agosto"
                              }
                              break;
                              case "08":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Septiembre"
                              }
                              break;
                              case "09":
                              if (dia > 30) {
                                dia = "01";
                                mes = "Octubre"
                              }
                              break;
                              case "10":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Noviembre"
                              }
                              break;
                              case "11":
                              if (dia > 30) {
                                dia = "01";
                                mes = "Diciembre"
                              }
                              break;
                              case "12":
                              if (dia > 31) {
                                dia = "01";
                                mes = "Enero"
                                anio = fecha.getFullYear() + 1;
                              }
                              break;
                          }
                          ciudadFecha = sheetData[2][4] + ", " + sheetData[1][4] + " a " + dia  + " de " + mes + " de " + anio;
                          fechaConstancia = dia  + " de " + mes + " de " + anio;
                          filename = anio.toString().substr(-2) + mesNumber + dia + "_Constancias Grupales - Búsqueda y Rescate - " + nombreComercial
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
                if (sheetData[1][4] == "Jalisco") {
                  if(sheetData[3][1] == "Evacuación, Búsqueda y Rescate"){
                    nombreCurso = "Evacuación de Inmuebles"
                    filename = fecha.getFullYear().toString().substr(-2) + mesNumber + dia + "_Constancias Grupales - Evacuación - " + nombreComercial
                    for (var i = 0; i < 2; i++) {
                      if(nombreCurso == "Evacuación de Inmuebles"){
                        generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantesEI, mult, numC);
                      }
                      else {
                        generatePDFGrupal(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets, fechaConstancia, nombreCurso, participantesBR, mult, numC);
                      }
                      nombreCurso = "Búsqueda y Rescate"
                      var anio = fecha.getFullYear();
                      dia = (fecha.getDate() >= 10) ? (fecha.getDate() + 1) : ("0" + (fecha.getDate() + 1))
                      switch(mesNumber) {
                        case "01":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Febrero"
                          }
                          break;
                          case "02":
                            var bisiesto = false;
                            var aniosRestantes = fecha.getFullYear();
                            do {
                              if (aniosRestantes == 2000) {
                                bisiesto = true;
                              }
                              aniosRestantes = aniosRestantes - 4;
                            }
                            while (aniosRestantes >= 2000)
                            if (bisiesto == true) {
                              if (dia > 29) {
                                dia = "01";
                                mes = "Marzo"
                              }
                            }
                            else {
                              if (dia > 28) {
                                dia = "01";
                                mes = "Marzo"
                              }
                            }
                          break;
                          case "03":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Abril"
                          }
                          break;
                          case "04":
                          if (dia > 30) {
                            dia = "01";
                            mes = "Mayo"
                          }
                          break;
                          case "05":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Junio"
                          }
                          break;
                          case "06":
                          if (dia > 30) {
                            dia = "01";
                            mes = "Julio"
                          }
                          break;
                          case "07":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Agosto"
                          }
                          break;
                          case "08":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Septiembre"
                          }
                          break;
                          case "09":
                          if (dia > 30) {
                            dia = "01";
                            mes = "Octubre"
                          }
                          break;
                          case "10":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Noviembre"
                          }
                          break;
                          case "11":
                          if (dia > 30) {
                            dia = "01";
                            mes = "Diciembre"
                          }
                          break;
                          case "12":
                          if (dia > 31) {
                            dia = "01";
                            mes = "Enero"
                            anio = fecha.getFullYear() + 1;
                          }
                          break;
                      }
                      ciudadFecha = sheetData[2][4] + ", " + sheetData[1][4] + " a " + dia  + " de " + mes + " de " + anio;
                      fechaConstancia = dia  + " de " + mes + " de " + anio;
                      filename = anio.toString().substr(-2) + mesNumber + dia + "_Constancias Grupales - Búsqueda y Rescate - " + nombreComercial
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
          //else {
          //  if (sheets[i - 1].name == "Registros PC") {
          //    readXlsxFile(file, { sheet: i }).then(function(data) {
          //      registrosData = data;
          //    })
          //  }
          //}
        }
      })
      document.getElementById('imgLogo').src = './img/404-error.png'; 
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
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
      document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px";  
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
    document.getElementById('imgLogo').style.width = "230px";
      document.getElementById('imgLogo').style.height = "200px"; 
    document.getElementById('input').files = null; 
    document.getElementById('txtFile').innerHTML = ''; 
    document.getElementById('txtFile').style.height = '25px';
  }
})

var generatePDFGrupal = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength, sheets, fechaConstancia, nombreCurso, participantes, mult, numC) => {
  try {
    if (sheetData[7][1] != null)
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
          doc.addImage(firmaALV, "JPEG", 227, 564, 175, 110)
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
        var direccion = sheetData[5][1].replace(/\n/g, ' ')
        var info = "Hago constar que el personal que labora en " + razonSocial[1] + " (" + nombreComercial +"), ubicado en " + direccion + " "
        + sheetData[2][4] + ", " + sheetData[1][4] + ", participó de manera satisfactoria en el *CURSO BASICO DE " + nombreCurso.toUpperCase() 
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
              //console.log(palabra + ": " + doc.getTextDimensions(palabra, {fontSize: 9.0}).w)
              //console.log(doc.getTextWidth(palabra))
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
                startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2 - (letrasMayusculas * 0.6);
                //startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2;
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
                //startX += doc.getTextDimensions(palabra, {fontSize: 9.0}).w + 2;
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
            //console.log(mytext + ": " + doc.getTextDimensions(mytext, {fontSize: 9.0}).w)
            
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
        doc.text(sheetData[3][4], 612/2, 670, 'center')
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
    
    if (sheetData[7][1] != null)
    {
      if (mult == true) {
        doc.save(filename + '_' + numC  + '.pdf')
      }
      else {
        doc.save(filename +'.pdf')
      }
    }
    if (finished >= sheetsLength - 2)
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
      console.log(error)
      document.getElementById("alertError").innerHTML='Algo paso, vuelvalo a intentar'
      loaderIcon.style.display = "none"
      loader.style.display = "block"
      displayError.style.display = "block"
    }
}

var generatePDFDC3 = (razonSocial, sheetData, filename, sheetsLength, sheets, date, nombreCurso, jaliscoCon) => {
  try {
    doc = null;
    var special = "";
    if (jaliscoCon == true) {
      switch(nombreCurso) {
        case "Evacuación de Inmuebles":
          special = "CEVA"
          break;
        case "Busqueda y Rescate":
          special = "CBYR"
          break;
      }
    }
    for (var a = 7; a < sheetData.length; a++)
    {
      if (special == ""){
        if (sheetData[a][1] != null)
      {
        if (doc == null) {
          var doc = new jsPDF('p', 'pt', [612, 792]);
          doc.addFileToVFS("calibril-normal.ttf", font_normal)
          doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
          doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
          doc.addFileToVFS("calibril-bold.ttf", font_bold)
          doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
          doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
        }
        else {
          doc.addPage(612, 792) //19.05, 25.4
        }

        doc.setFont("calibri-bold");
        doc.setFontSize(10);
        doc.text("FORMATO DC-3", 306, 113, "center")
        doc.text("CONSTANCIA DE COMPETENCIAS O DE HABILIDADES LABORALES", 306, 125, "center")
        doc.setTextColor( 255, 255, 255 )
        var suburbiaChecked= false;
        if (firmaElectronica.checked == true) {
          doc.addImage(firmaALV, "JPEG", 64, 514, 175, 110)
          if (razonSocial[1].toUpperCase().includes("SUBURBIA")) {
            suburbiaChecked = true;
            if (sheetData[4][6] != null && sheetData[4][6] != '' && (sheetData[4][6].toUpperCase() == 'JOSÉ ESCOTO GARCÍA'  || sheetData[4][6].toUpperCase() == 'JOSE ESCOTO GARCIA')) {
              doc.addImage(firmaREP1, "JPEG", 254, 492, 110, 120)
            }
            if (sheetData[5][6] != null && sheetData[5][6] != '' && sheetData[5][6].toUpperCase() == 'ADRIANA RIVERA GUIJOSA') {
              doc.addImage(firmaREP2, "JPEG", 429, 489, 70, 140)
            }
          }
        }
        doc.addImage(imgLavi, "JPEG", 420.743778, 48.84646, 154.256222, 47.15354); // 0.5, 0.5, 9, 2.7
        doc.addImage(logoClient, "JPEG", 37, 48.84646, 154.256222, 47.15354);    
        doc.setLineWidth(1.5);
        doc.rect(37, 140, 538, 18, 'F'); // filled square 
        doc.rect(37, 262, 538, 18, 'F');
        doc.rect(37, 354, 538, 18, 'F');
        doc.text("DATOS DEL TRABAJADOR", 306, 152, "center") 
        doc.text("DATOS DE LA EMPRESA", 306, 274, "center")
        doc.setFontSize(9);
        doc.text("DATOS DEL PROGRAMA DE CAPACITACIÓN, ADIESTRAMIENTO Y PRODUCTIVIDAD", 306, 366, "center")

        // Datos del trabajador
        doc.setTextColor( 0, 0, 0 )
        doc.rect(37.5, 158, 537, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre (Anotar apellido paterno, apellido materno y nombre)", 38.5, 168)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(sheetData[a][10], 39.5, 184)
        doc.rect(37.5, 190, 310.5, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Clave Única de Registro de Población", 38.5, 200)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        if (sheetData[a][4] != null) {
          doc.text(sheetData[a][4].toUpperCase(), 39.5, 216)
        }
        
        doc.rect(348, 190, 226.5, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Ocupación específica (Catálogo Nacional de Ocupaciones) ₁", 349, 200)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("07.1 Comercio", 350, 216)
        doc.rect(37.5, 222, 537, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Puesto*", 38.5, 232)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        if (sheetData[a][5] != null) {
          doc.text(sheetData[a][5].toUpperCase(), 39.5, 248)
        }
        

        doc.setDrawColor(0, 0, 0);
        
        //Datos de la empresa
        doc.rect(37.5, 280, 537, 64); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre o razón social (En caso de persona física, anotar apellido paterno, apellido materno y nombre(s))", 38.5, 290)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(razonSocial[1].toUpperCase(), 306, 306, "center")
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Registro Federal de Contribuyentes (SHCP)", 38.5, 322)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(sheetData[3][6].toUpperCase(), 306, 338, "center")

        // DATOS DEL PROGRAMA DE CAPACITACIÓN, ADIESTRAMIENTO Y PRODUCTIVIDAD
        doc.rect(37.5, 372, 537, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre del curso", 38.5, 380)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(nombreCurso, 39.5, 391)
        doc.rect(37.5, 394, 140.5, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Duración en horas", 38.5, 402)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        for (var b = 0; b < duracion.length; b++){
          if (nombreCurso == duracion[b].curso){
            doc.text((duracion[b].duracion > 1) ? (duracion[b].duracion + " horas") : (duracion[b].duracion +" hora"), 39.5, 413)
          }
        }
        doc.rect(178, 394, 102, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Periodo de ejecición:", 187, 406)
        doc.text("De", 265, 412)
        doc.rect(280, 394, 69.25, 22); // AÑO 1
        doc.setFontSize(8.5);
        doc.text("Año", 314.625, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(0, 1), 287.15625, 412.7);
        doc.text(date.substr(1, 1), 304.46875, 412.7);
        doc.text(date.substr(2, 1), 321.78125, 412.7);
        doc.text(date.substr(3, 1), 339.09375, 412.7);
        // 17.3125
        doc.line(297.3125, 405, 297.3125, 416);
        doc.line(314.625, 405, 314.625, 416);
        doc.line(331.9375, 405, 331.9375, 416);

        doc.rect(349.25, 394, 34.25, 22); // MES 1
        doc.setFont("calibri-normal");
        doc.text("Mes", 366.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(5, 1), 356.40625, 412.7);
        doc.text(date.substr(6, 1), 373.71875, 412.7);
        // 17.125
        doc.line(366.375, 405, 366.375, 416);

        doc.rect(383.5, 394, 34.25, 22); // DIA 1
        doc.setFont("calibri-normal");
        doc.text("Dia", 400.625, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(8, 1), 391.03125, 412.7);
        doc.text(date.substr(9, 1), 408.34375, 412.7);
        // 17.125
        doc.line(400.625, 405, 400.625, 416);

        doc.rect(417.75, 394, 19, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("a", 425.25, 412)

        doc.rect(436.75, 394, 69.25, 22); // AÑO 2
        doc.setFontSize(8.5);
        doc.text("Año", 471.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(0, 1), 443.90625, 412.7);
        doc.text(date.substr(1, 1), 461.21875, 412.7);
        doc.text(date.substr(2, 1), 478.53125, 412.7);
        doc.text(date.substr(3, 1), 495.84375, 412.7);
        // 17.3125
        doc.line(454.0625, 405, 454.0625, 416);
        doc.line(471.375, 405, 471.375, 416);
        doc.line(488.6875, 405, 488.6875, 416);

        doc.rect(506, 394, 34.25, 22); // MES 2
        doc.setFont("calibri-normal");
        doc.text("Mes", 523.125, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(5, 1), 513.15625, 412.7);
        doc.text(date.substr(6, 1), 530.46875, 412.7);
        // 17.125
        doc.line(523.125, 405, 523.125, 416);

        doc.rect(540.25, 394, 34.25, 22); // DIA 2
        doc.setFont("calibri-normal");
        doc.text("Dia", 557.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(8, 1), 547.78125, 412.7);
        doc.text(date.substr(9, 1), 565.09375, 412.7);
        // 17.125
        doc.line(557.375, 405, 557.375, 416);


        doc.rect(37.5, 416, 537, 29); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Área temática del curso ₂", 38.5, 424)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("6000 - Seguridad", 39.5, 440)
        doc.rect(37.5, 445, 537, 29); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre del agente capacitador o STPS ₃", 38.5, 456)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("LAVI Fire Workshop México, S.A. de C.V. (LFW-160516-NJ1-0013)", 39.5, 472)
        //574.5

        //Espacio de firmas
        doc.rect(37.5, 489, 537, 121.6); 
        doc.setFont("calibri-bold");
        doc.setFontSize(7);
        doc.text("Los datos se asientan en esta constancia bajo protesta de decir la verdad, apercibidos de la responsabilidad en que incurre todo aquel que no se", 306, 499, "center")
        doc.text("conduce con la verdad", 306, 509, "center")
        doc.setFont("calibri-normal");
        doc.text("Instructor o tutor", 148, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        doc.text("ANTONIO LAVIN VILLA", 148, 589, "center")
        doc.line(79, 592, 217, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 148, 599, "center")
        doc.setFont("calibri-normal");
        doc.text("Patrón o representante legal ₄", 306, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        if (sheetData[4][6] != null && sheetData[4][6] != '') {
          doc.text(sheetData[4][6].toUpperCase(), 306, 589, "center")
        }
        doc.line(237, 592, 375, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 306, 599, "center")
        doc.setFont("calibri-normal");
        doc.text("Representante de los trabajadores ₅", 464, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        if (sheetData[5][6] != null && sheetData[5][6] != '') {
          doc.text(sheetData[5][6].toUpperCase(), 464, 589, "center")
        }
        doc.line(395, 592, 533, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 464, 599, "center")


        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        doc.text("INSTRUCCIONES", 38.5, 643.6)
        doc.setFont("calibri-normal")
        doc.setFontSize(6.5);
        doc.text("- Llenar a máquina o con letra de molde", 43, 655.6)
        doc.text("- Deberá entregarse al trabajador dentro de los veinte días hábiles siguientes al término del curso de capacitación aprobado.", 43, 664.6)
        doc.text("1. Las áreas y subáreas ocupacionales del Catálogo Nacional de Ocupaciones se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx", 39, 673.6)
        doc.text("2. Las áreas temáticas de los cursos se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx", 39, 682.6)
        doc.text("3. Cursos impartidos por el área competente de la Secretaria del Trabajo y Previsión Social.", 39, 691.6)
        doc.text("4. Para empresas con menos de 51 trabajadores. Para empresas con más de 50 trabajadores firmaría el representante del patrón ante la Comisión mixta de capacitación, adiestramiento y productividad.", 39, 700.6)
        doc.text("5. Solo para empresascon más de 50 trabajadores.", 39, 709.6)
        doc.text("*Dato no obligatorio.", 39, 718.6)
        doc.setFontSize(7.5);
        doc.text("DC-3", 534.5, 723.6)
        doc.text("ANVERSO", 520.5, 732.6)
      }
      }
      else {
        if (sheetData[a][1] != null && sheetData[a][11].includes(special) == true && sheetData[a][11].includes("Agregado") == false)
      {
        if (doc == null) {
          var doc = new jsPDF('p', 'pt', [612, 792]);
          doc.addFileToVFS("calibril-normal.ttf", font_normal)
          doc.addFont("calibril-normal.ttf", "calibri-normal", "normal");
          doc.addFont("calibril-normal.ttf", "calibri-normal", "bold");
          doc.addFileToVFS("calibril-bold.ttf", font_bold)
          doc.addFont("calibril-bold.ttf", "calibri-bold", "normal");
          doc.addFont("calibril-bold.ttf", "calibri-bold", "bold");
        }
        else {
          doc.addPage(612, 792) //19.05, 25.4
        }

        doc.setFont("calibri-bold");
        doc.setFontSize(10);
        doc.text("FORMATO DC-3", 306, 113, "center")
        doc.text("CONSTANCIA DE COMPETENCIAS O DE HABILIDADES LABORALES", 306, 125, "center")
        doc.setTextColor( 255, 255, 255 )
        var suburbiaChecked= false;
        if (firmaElectronica.checked == true) {
          doc.addImage(firmaALV, "JPEG", 64, 514, 175, 110)
          if (razonSocial[1].toUpperCase().includes("SUBURBIA")) {
            suburbiaChecked = true;
            if (sheetData[4][6] != null && sheetData[4][6] != '' && (sheetData[4][6].toUpperCase() == 'JOSÉ ESCOTO GARCÍA'  || sheetData[4][6].toUpperCase() == 'JOSE ESCOTO GARCIA')) {
              doc.addImage(firmaREP1, "JPEG", 254, 492, 110, 120)
            }
            if (sheetData[5][6] != null && sheetData[5][6] != '' && sheetData[5][6].toUpperCase() == 'ADRIANA RIVERA GUIJOSA') {
              doc.addImage(firmaREP2, "JPEG", 429, 489, 70, 140)
            }
          }
        }
        doc.addImage(imgLavi, "JPEG", 420.743778, 48.84646, 154.256222, 47.15354); // 0.5, 0.5, 9, 2.7
        doc.addImage(logoClient, "JPEG", 37, 48.84646, 154.256222, 47.15354);    
        doc.setLineWidth(1.5);
        doc.rect(37, 140, 538, 18, 'F'); // filled square 
        doc.rect(37, 262, 538, 18, 'F');
        doc.rect(37, 354, 538, 18, 'F');
        doc.text("DATOS DEL TRABAJADOR", 306, 152, "center") 
        doc.text("DATOS DE LA EMPRESA", 306, 274, "center")
        doc.setFontSize(9);
        doc.text("DATOS DEL PROGRAMA DE CAPACITACIÓN, ADIESTRAMIENTO Y PRODUCTIVIDAD", 306, 366, "center")

        // Datos del trabajador
        doc.setTextColor( 0, 0, 0 )
        doc.rect(37.5, 158, 537, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre (Anotar apellido paterno, apellido materno y nombre)", 38.5, 168)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(sheetData[a][10], 39.5, 184)
        doc.rect(37.5, 190, 310.5, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Clave Única de Registro de Población", 38.5, 200)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        if (sheetData[a][4] != null) {
          doc.text(sheetData[a][4].toUpperCase(), 39.5, 216)
        }
        
        doc.rect(348, 190, 226.5, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Ocupación específica (Catálogo Nacional de Ocupaciones) ₁", 349, 200)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("07.1 Comercio", 350, 216)
        doc.rect(37.5, 222, 537, 32); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Puesto*", 38.5, 232)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        if (sheetData[a][5] != null) {
          doc.text(sheetData[a][5].toUpperCase(), 39.5, 248)
        }
        

        doc.setDrawColor(0, 0, 0);
        
        //Datos de la empresa
        doc.rect(37.5, 280, 537, 64); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre o razón social (En caso de persona física, anotar apellido paterno, apellido materno y nombre(s))", 38.5, 290)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(razonSocial[1].toUpperCase(), 306, 306, "center")
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Registro Federal de Contribuyentes (SHCP)", 38.5, 322)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(sheetData[3][6].toUpperCase(), 306, 338, "center")

        // DATOS DEL PROGRAMA DE CAPACITACIÓN, ADIESTRAMIENTO Y PRODUCTIVIDAD
        doc.rect(37.5, 372, 537, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre del curso", 38.5, 380)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text(nombreCurso, 39.5, 391)
        doc.rect(37.5, 394, 140.5, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Duración en horas", 38.5, 402)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        for (var b = 0; b < duracion.length; b++){
          if (nombreCurso == duracion[b].curso){
            doc.text((duracion[b].duracion > 1) ? (duracion[b].duracion + " horas") : (duracion[b].duracion +" hora"), 39.5, 413)
          }
        }
        doc.rect(178, 394, 102, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Periodo de ejecición:", 187, 406)
        doc.text("De", 265, 412)
        doc.rect(280, 394, 69.25, 22); // AÑO 1
        doc.setFontSize(8.5);
        doc.text("Año", 314.625, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(0, 1), 287.15625, 412.7);
        doc.text(date.substr(1, 1), 304.46875, 412.7);
        doc.text(date.substr(2, 1), 321.78125, 412.7);
        doc.text(date.substr(3, 1), 339.09375, 412.7);
        // 17.3125
        doc.line(297.3125, 405, 297.3125, 416);
        doc.line(314.625, 405, 314.625, 416);
        doc.line(331.9375, 405, 331.9375, 416);

        doc.rect(349.25, 394, 34.25, 22); // MES 1
        doc.setFont("calibri-normal");
        doc.text("Mes", 366.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(5, 1), 356.40625, 412.7);
        doc.text(date.substr(6, 1), 373.71875, 412.7);
        // 17.125
        doc.line(366.375, 405, 366.375, 416);

        doc.rect(383.5, 394, 34.25, 22); // DIA 1
        doc.setFont("calibri-normal");
        doc.text("Dia", 400.625, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(8, 1), 391.03125, 412.7);
        doc.text(date.substr(9, 1), 408.34375, 412.7);
        // 17.125
        doc.line(400.625, 405, 400.625, 416);

        doc.rect(417.75, 394, 19, 22); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("a", 425.25, 412)

        doc.rect(436.75, 394, 69.25, 22); // AÑO 2
        doc.setFontSize(8.5);
        doc.text("Año", 471.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(0, 1), 443.90625, 412.7);
        doc.text(date.substr(1, 1), 461.21875, 412.7);
        doc.text(date.substr(2, 1), 478.53125, 412.7);
        doc.text(date.substr(3, 1), 495.84375, 412.7);
        // 17.3125
        doc.line(454.0625, 405, 454.0625, 416);
        doc.line(471.375, 405, 471.375, 416);
        doc.line(488.6875, 405, 488.6875, 416);

        doc.rect(506, 394, 34.25, 22); // MES 2
        doc.setFont("calibri-normal");
        doc.text("Mes", 523.125, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(5, 1), 513.15625, 412.7);
        doc.text(date.substr(6, 1), 530.46875, 412.7);
        // 17.125
        doc.line(523.125, 405, 523.125, 416);

        doc.rect(540.25, 394, 34.25, 22); // DIA 2
        doc.setFont("calibri-normal");
        doc.text("Dia", 557.375, 402, "center")
        doc.setFont("calibri-bold");
        doc.text(date.substr(8, 1), 547.78125, 412.7);
        doc.text(date.substr(9, 1), 565.09375, 412.7);
        // 17.125
        doc.line(557.375, 405, 557.375, 416);


        doc.rect(37.5, 416, 537, 29); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Área temática del curso ₂", 38.5, 424)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("6000 - Seguridad", 39.5, 440)
        doc.rect(37.5, 445, 537, 29); // empty square
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre del agente capacitador o STPS ₃", 38.5, 456)
        doc.setFont("calibri-bold");
        doc.setFontSize(8.5);
        doc.text("LAVI Fire Workshop México, S.A. de C.V. (LFW-160516-NJ1-0013)", 39.5, 472)
        //574.5

        //Espacio de firmas
        doc.rect(37.5, 489, 537, 121.6); 
        doc.setFont("calibri-bold");
        doc.setFontSize(7);
        doc.text("Los datos se asientan en esta constancia bajo protesta de decir la verdad, apercibidos de la responsabilidad en que incurre todo aquel que no se", 306, 499, "center")
        doc.text("conduce con la verdad", 306, 509, "center")
        doc.setFont("calibri-normal");
        doc.text("Instructor o tutor", 148, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        doc.text("ANTONIO LAVIN VILLA", 148, 589, "center")
        doc.line(79, 592, 217, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 148, 599, "center")
        doc.setFont("calibri-normal");
        doc.text("Patrón o representante legal ₄", 306, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        if (sheetData[4][6] != null && sheetData[4][6] != '') {
          doc.text(sheetData[4][6].toUpperCase(), 306, 589, "center")
        }
        doc.line(237, 592, 375, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 306, 599, "center")
        doc.setFont("calibri-normal");
        doc.text("Representante de los trabajadores ₅", 464, 539, "center")
        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        if (sheetData[5][6] != null && sheetData[5][6] != '') {
          doc.text(sheetData[5][6].toUpperCase(), 464, 589, "center")
        }
        doc.line(395, 592, 533, 592);
        doc.setFont("calibri-normal");
        doc.setFontSize(7);
        doc.text("Nombre y firma", 464, 599, "center")


        doc.setFont("calibri-bold");
        doc.setFontSize(9);
        doc.text("INSTRUCCIONES", 38.5, 643.6)
        doc.setFont("calibri-normal")
        doc.setFontSize(6.5);
        doc.text("- Llenar a máquina o con letra de molde", 43, 655.6)
        doc.text("- Deberá entregarse al trabajador dentro de los veinte días hábiles siguientes al término del curso de capacitación aprobado.", 43, 664.6)
        doc.text("1. Las áreas y subáreas ocupacionales del Catálogo Nacional de Ocupaciones se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx", 39, 673.6)
        doc.text("2. Las áreas temáticas de los cursos se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx", 39, 682.6)
        doc.text("3. Cursos impartidos por el área competente de la Secretaria del Trabajo y Previsión Social.", 39, 691.6)
        doc.text("4. Para empresas con menos de 51 trabajadores. Para empresas con más de 50 trabajadores firmaría el representante del patrón ante la Comisión mixta de capacitación, adiestramiento y productividad.", 39, 700.6)
        doc.text("5. Solo para empresascon más de 50 trabajadores.", 39, 709.6)
        doc.text("*Dato no obligatorio.", 39, 718.6)
        doc.setFontSize(7.5);
        doc.text("DC-3", 534.5, 723.6)
        doc.text("ANVERSO", 520.5, 732.6)
      }
      }
      
    }
    if (doc != null)
    {
      doc.save(filename + '.pdf')
      doc = null;
    }
    if (finished >= sheetsLength - 2)
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
      console.log(error)
    }
}