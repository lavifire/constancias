import { font_normal } from '../css/calibril-normal.js';
import { font_bold } from '../css/calibril-bold.js';
var inputFile = document.getElementById("input");
var inputLogo = document.getElementById("inputLogo");
var imgLogo = document.getElementById("imgLogo");
var imgLavi = new Image()
imgLavi.src = "./img/LAVI Fire_HI-RES.jpg"
var imgEsquina = new Image()
var RegistrosData;
imgEsquina.src = "./img/esquina/esquina-rojo.png"
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
var codigos = [""];
var tipo = "";
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

inputFile.addEventListener("change", function() {
  //if (codigos[0] != ""){
    console.log(inputFile.files[0])
    loader.style.display = "block"
    readXlsxFile(inputFile.files[0], { getSheets: true }).then(function(sheets) {
      console.log(sheets)
      for (var i = sheets.length; i > 0; i--) {
        if (sheets[i - 1].name != "Registros PC" && sheets[i - 1].name != "Estados Municipios") {
          readXlsxFile(inputFile.files[0], { sheet: i }).then(function(sheetData) {
            console.log(sheetData)
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

            for (var b = 3; b < RegistrosData.length; b++)
            {
              if (codigos[0] == RegistrosData[b][4])
              {
                tipo = RegistrosData[b][2]
                console.log(tipo)
              }
            }

            var dia = (fecha.getDate() >= 10) ? (fecha.getDate()) : ("0" + fecha.getDate())
            var ciudadFecha = sheetData[1][4] + ", " + sheetData[0][4] + " a " + dia  + " de " + mes + " de " + fecha.getFullYear();
            console.log(ciudadFecha)
            console.log(codigos);
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
            generatePDF(ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheets.length - 2, sheets);
          })
        }
        else {
          if (sheets[i - 1].name == "Registros PC")
          {
            readXlsxFile(inputFile.files[0], { sheet: i }).then(function(sheetData) {
              RegistrosData = sheetData;
            })
            wait(1500)
          }
        }
      }
    })
    txtFile.innerHTML = inputFile.files[0].name;
    txtFile.style.height = "auto"
  //}
  //else {
  //  alert("No se ha elegido un estado");
  //}
  
})

var generatePDF = (ciudadFecha, razonSocial, nombreComercial, sheetData, filename, sheetsLength, sheets) => {
  console.log(filename)
  for (var a = 5; a < sheetData.length; a++)
  {
    if (a == 5)
    {
      if (sheetData[5][2] != null)
      {
        var doc = new jsPDF('l', 'pt', [540, 720]) //19.05, 25.4
        console.log("Entro")
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
      
      doc.addImage(image, "png", 56.6929, 445.0394, 51.0236, 51.0236) // 2, 15.7, 1.8, 1.8
  
      doc.setFontSize(10);
      //doc.text(text, 720/2, 85.0394, 'center') // 25.4/2, 3
      doc.text(ciudadFecha, 651.969 - doc.getTextDimensions(ciudadFecha).w, 85.0394);
      doc.text("LAVI Fire Workshop México, S.A. de C.V.", 720/2, 322.2443, 'center'); // 25.4/2, 11
      doc.text(tipo, 720/2, 336.4176, 'center'); // 25.4/2, 11.5
      doc.text(codigos[0], 720/2, 350.5903, 'center'); // 25.4/2, 12
      if (codigos.length > 1)
      {
        doc.text(codigos[1], 720/2, 364.764, 'center'); // 25.4/2, 12.5
      }
      if (sheetData[1][4] == "León")
      {
        doc.text("Arq. Antonio Lavín Villa", 119.999055, 411.0236); // 4.2333, 14.5
        doc.text("Javier Everardo Hernández Moreno", 479.9990551, 411.0236); // 16.9333, 14.5
        doc.setFontSize(8.5);
        doc.text("Representante Legal", 131.337638, 425.197); // 4.6333, 15
        doc.text("Instructor del curso", 522.5187402, 425.197); // 18.4333, 15
      }
      else
      {
        doc.text("Arq. Antonio Lavín Villa", 720/2, 411.0236, 'center'); // 25.4/2, 14.5
        doc.setFontSize(8.5);
        doc.text("Representante Legal", 720/2, 425.197, 'center'); // 25.4/2, 15
      }
      doc.setFontSize(12);
      doc.text("Otorga la presente constancia a:", 720/2, 121.831, 'center'); // 25.4/2, 5
      doc.text("Por haber participado en el curso de:", 720/2, 198.425, 'center'); // 25.4/2, 7
      doc.text(nombreComercial, 720/2, 289.291, 'center') // 25.4/2, 9.5
      var startX = (720 / 2) - (doc.getTextDimensions(text).w / 2);
      const arrayOfNormalAndBoldText = txtInfo.split('**');
      arrayOfNormalAndBoldText.map((mytext, i) => {
      // every even item is a normal font weight item
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
      doc.setTextColor( 177, 177, 177 )
      doc.setFontSize(8.5);
      doc.text("info@lavi.com.mx", 720/2, 475, 'center'); // 360
      doc.text("O. (477).713.0016", 490, 470, 'center');
      doc.text("M. (477).670.2737", 490, 480, 'center');
      doc.text("Av. Las Américas 401-B", 230, 465, 'center');
      doc.text("Col. Andrade CP 37020", 230, 475, 'center');
      doc.text("León, Guanajuato", 230, 485, 'center');
      //doc.fromHTML(htmlinfo.innerHTML, 5, 8.5) 
    }
  }
  if (sheetData[5][2] != null)
  {
    doc.save(filename + '.pdf')
  }
    
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

inputLogo.addEventListener("change", () => {
  var reader = new FileReader();
  reader.onloadend = function() {
    imgLogo.src = reader.result;
    imgLogo.style.height = "100px";
    imgLogo.style.width = "125px";
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