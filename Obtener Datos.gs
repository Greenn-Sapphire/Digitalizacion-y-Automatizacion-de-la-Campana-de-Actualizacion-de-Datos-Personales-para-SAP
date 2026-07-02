const Libro = SpreadsheetApp.getActiveSpreadsheet();
let Nombre_Hoja;
let Nombre_Hoja_Mostrar;

ID_Libro_Destino = "1Up_MnEBaEYvdTLLNuJp2dRZSIheKq0jfh40O0BG1OcU"

//Variables Infotipo 06 - Sección 1 Datos Generales
let Numero_Colaborador;
let Nombre_Colaborador;
let Area;
let Puesto;
let Direccion_Registrada;
let Fecha_de_Nacimiento;
let Telefono_Casa;
let Telefono_Celular;
let Ubicacion;
let Direccion_Principal_Separada;
let Direccion_Emergencia_Separada;

//Variables Infotipo 21 - Sección 2 Datos Familiares
let Estado_Civil;
let Nombre_Pareja;
let Apellido_Pareja;
let Segundo_Apellido_Pareja;
let Sexo_Pareja;
let Fecha_de_Nacimiento_Pareja;
let Edad_Pareja;
let Grado_Escolar_Pareja;

let Familiares = [];

let Celular_Trabajo;
let Correo_Trabajo;

//Variables Infotipo 06 - Sección 3 Datos Contacto de Emergencia
let Nombre_Contacto_Emergencia;
let Parentesco;
let Direccion;
let Tel_Celular_1;
let Tel_Celular_2;
let Tel_Casa;

let Camisa_Admon;
let Camisa_Admon_Manga_Larga;
let Pantalon_Admon;
let Botas_Admon;

let Camisola;
let Pantalon_Ope;
let Botas;
let Playera_Blanca;

let Playera_Casual;
let Chamarra_Chaleco;

let Datos_Infotipo_IT0021 = [];

function get_Data(Usuario, Accion) {
  switch(Usuario){
    case 'plovera@aguakan.com':
      if(Accion === "1"){
        Nombre_Hoja = 'Vista Colaborador Pablo';
        Nombre_Hoja_Mostrar = 'Edición Pablo';
      }else{
        Nombre_Hoja = 'Edición Pablo';
      }
      break;
    case 'ipech@aguakan.com':
      if(Accion === "1"){
        Nombre_Hoja = 'Vista Colaborador Irving';
        Nombre_Hoja_Mostrar = 'Edición Irving';
      }else{
        Nombre_Hoja = 'Edición Irving';
      }
      break;
    case 'bduran@aguakan.com':
    case 'ntec@aguakan.com':
      if(Accion === "1"){
        Nombre_Hoja = 'Vista Colaborador Nadia';
        Nombre_Hoja_Mostrar = 'Edición Nadia';
      }else{
        Nombre_Hoja = 'Edición Nadia';
      }
      break;
    case 'kvillamil@aguakan.com':
      if(Accion === "1"){
        Nombre_Hoja = 'Vista Colaborador Karina';
        Nombre_Hoja_Mostrar = 'Edición Karina';
      }else{
        Nombre_Hoja = 'Edición Karina';
      }
      break;
    default:
      SpreadsheetApp.getUi().alert('Usuario no autorizado');
      return;
  }

  const Copiar_Hoja = Libro.getSheetByName(Nombre_Hoja);
  const Pegar_Copia = Libro.getSheetByName(Nombre_Hoja_Mostrar);

  Nombre_Colaborador = Copiar_Hoja.getRange("B6").getValue();
  Numero_Colaborador = Copiar_Hoja.getRange("K6").getValue();

  Area = Copiar_Hoja.getRange("B7").getValue();
  Fecha_de_Nacimiento = Copiar_Hoja.getRange("K7").getValue();

  Puesto = Copiar_Hoja.getRange("B8").getValue();
  Telefono_Casa = Copiar_Hoja.getRange("K8").getValue();
  Telefono_Celular = Copiar_Hoja.getRange("K9").getValue();

  Direccion_Principal_Separada = [
    Copiar_Hoja.getRange("B10").getValue(), //Calle y número
    Copiar_Hoja.getRange("D10").getValue(), //No. Casa
    Copiar_Hoja.getRange("E10").getValue(), //Distrito
    Copiar_Hoja.getRange("G10").getValue(), //Población
    Copiar_Hoja.getRange("H10").getValue(), //Código de la ciudad
    Copiar_Hoja.getRange("I10").getValue().split(",").map(p => p.trim())[0] ?? "", //Región
    Copiar_Hoja.getRange("I10").getValue().split(",").map(p => p.trim())[1] ?? "", //Clave de país
    Copiar_Hoja.getRange("J10").getValue()  //Código post./ Pobl.
  ]
  Ubicacion = Copiar_Hoja.getRange("L10").getValue();

  Estado_Civil = Copiar_Hoja.getRange("B13").getValue();

  Nombre_Pareja = Copiar_Hoja.getRange("B16").getValue();
  Apellido_Pareja = Copiar_Hoja.getRange("D16").getValue();
  Segundo_Apellido_Pareja = Copiar_Hoja.getRange("F16").getValue();
  Sexo_Pareja = Copiar_Hoja.getRange("H16").getValue();
  Fecha_de_Nacimiento_Pareja = Copiar_Hoja.getRange("I16").getValue();
  Edad_Pareja = Copiar_Hoja.getRange("K16").getValue();
  Grado_Escolar_Pareja = Copiar_Hoja.getRange("L16").getValue();

  for (let i = 17; i <= 22; i++) {
    let Sexo = Copiar_Hoja.getRange(`H${i}`).getValue();
    if (Sexo === "H" || Sexo === "M") {
      Familiares.push([
        Copiar_Hoja.getRange(`B${i}`).getValue(), // Nombre 0
        "",                                         // 1
        Copiar_Hoja.getRange(`D${i}`).getValue(), // Apellido 2
        "",                                         // 3
        Copiar_Hoja.getRange(`F${i}`).getValue(), // Segundo apellido 4
        "",                                         // 5
        Sexo,                                    // Sexo 6
        Copiar_Hoja.getRange(`I${i}`).getValue(), // Fecha de nacimiento 7
        "",                                         // 8
        Copiar_Hoja.getRange(`K${i}`).getValue(), // Edad 9
        Copiar_Hoja.getRange(`L${i}`).getValue()  // Grado escolar 10
      ]);
    }else{
      Familiares.push(["", "", "", "", "", "", "", "", "", "", ""])
    }
  }

  Nombre_Contacto_Emergencia = Copiar_Hoja.getRange("B26").getValue();
  Parentesco = Copiar_Hoja.getRange("B27").getValue();
  Direccion_Emergencia_Separada = [
    Copiar_Hoja.getRange("B28").getValue(), //Calle y número 0
    Copiar_Hoja.getRange("C28").getValue(), //No. Casa 1
    Copiar_Hoja.getRange("D28").getValue(), //Distrito 2
    Copiar_Hoja.getRange("E28").getValue(), //Población 3
    Copiar_Hoja.getRange("F28").getValue(), //Código de la ciudad 4
    Copiar_Hoja.getRange("G28").getValue().split(",").map(p => p.trim())[0] ?? "", //Región 5
    Copiar_Hoja.getRange("G28").getValue().split(",").map(p => p.trim())[1] ?? "", //Clave de país 6
    Copiar_Hoja.getRange("H28").getValue()  //Código post./ Pobl. 7
  ]
  Tel_Celular_1 = Copiar_Hoja.getRange("D29").getValue();
  Tel_Celular_2 = Copiar_Hoja.getRange("D30").getValue();
  Tel_Casa = Copiar_Hoja.getRange("B31").getValue();

  Camisa_Admon = Copiar_Hoja.getRange("J26").getValue();
  Camisa_Admon_Manga_Larga = Copiar_Hoja.getRange("J27").getValue();
  Pantalon_Admon = Copiar_Hoja.getRange("J28").getValue();
  Botas_Admon = Copiar_Hoja.getRange("J29").getValue();

  Camisola = Copiar_Hoja.getRange("L26").getValue();
  Pantalon_Ope = Copiar_Hoja.getRange("L27").getValue();
  Botas = Copiar_Hoja.getRange("L28").getValue();
  Playera_Blanca = Copiar_Hoja.getRange("L29").getValue();

  Playera_Casual = Copiar_Hoja.getRange("J30").getValue();
  Chamarra_Chaleco = Copiar_Hoja.getRange("L30").getValue();
  
  Celular_Trabajo = Copiar_Hoja.getRange("J31").getValue();
  Correo_Trabajo = Copiar_Hoja.getRange("L31").getValue();

  if(Accion === "1"){
    show_Data(Pegar_Copia)
  }else{
    switch (Direccion_Principal_Separada[4]) {
      case "Isla Mujeres":
        Direccion_Principal_Separada[4] = 3;
        break;
      case "Benito Juarez":
        Direccion_Principal_Separada[4] = 5;
        break;
      case "Solidaridad":
        Direccion_Principal_Separada[4] = 8;
        break;
      case "Puerto Morelos":
        Direccion_Principal_Separada[4] = 11;
        break;
      case "Tizimin":
        Direccion_Principal_Separada[4] = 96;
        break;
      case "Valladolid":
        Direccion_Principal_Separada[4] = 102;
        break;
      case "Merida":
        Direccion_Principal_Separada[4] = 50;
        break;
      default:
        Direccion_Principal_Separada[4] = 1010;
    }

    switch (Direccion_Emergencia_Separada[4]) {
      case "Isla Mujeres":
        Direccion_Emergencia_Separada[4] = 3;
        break;
      case "Benito Juarez":
        Direccion_Emergencia_Separada[4] = 5;
        break;
      case "Solidaridad":
        Direccion_Emergencia_Separada[4] = 8;
        break;
      case "Puerto Morelos":
        Direccion_Emergencia_Separada[4] = 11;
        break;
      case "Tizimin":
        Direccion_Emergencia_Separada[4] = 96;
        break;
      case "Valladolid":
        Direccion_Emergencia_Separada[4] = 102;
        break;
      case "Merida":
        Direccion_Emergencia_Separada[4] = 50;
        break;
      default:
        Direccion_Emergencia_Separada[4] = 1010;
    }

    let Datos_Infotipo_IT0006 = [[
      Numero_Colaborador, //Número de personal
      "1", //Subtipo (Información del colaborador)
      "", //Identif.objeto
      "", //Indicador bloqueo
      "31/12/9999", //Fin de validez
      "", //Inicio de validez
      "0", //Nº regist.infotipo
      "1", //Clase de dirección
      "", //Nombre c/o
      Direccion_Principal_Separada[0], //Calle y número
      Direccion_Principal_Separada[1], //No. Casa
      Direccion_Principal_Separada[3], //Población
      Direccion_Principal_Separada[2], //Distrito
      Direccion_Principal_Separada[7], //Código post./ Pobl.
      Direccion_Principal_Separada[6], //Clave de país
      Telefono_Celular, //Nº teléfono
      "0", //Distancia en km
      "0", //Viviendas p.personal
      "", //Ruta del autobús
      "", //Cpo.adicional direc.
      "", //Calle 2
      "", //Calle 3
      Direccion_Principal_Separada[5], //Región
      "", //Nº (edificio)
      "", //N°Vivienda
      "", //Sigla del edificio
      "", //Piso
      "", //Sigla vía pública
      "0", //Distancia en km
      "TEL2", //Clase comunicación
      Telefono_Casa, //Número
      "", //Clase comunicación
      "", //Número
      Direccion_Principal_Separada[4], //Código de la ciudad°
      "0" //Subscrip.soc.ferrocarril
    ]];
    if(Nombre_Contacto_Emergencia){
      Datos_Infotipo_IT0006.push([
        Numero_Colaborador, //Número de personal
        "4", //Subtipo (Información del contacto de emergencia)
        "", //Identif.objeto
        "", //Indicador bloqueo
        "31/12/9999", //Fin de validez
        "", //Inicio de validez
        "0", //Nº regist.infotipo
        "4", //Clase de dirección
        Nombre_Contacto_Emergencia + ", " + Parentesco, //Nombre c/o
        Direccion_Emergencia_Separada[0], //Calle y número
        Direccion_Emergencia_Separada[1], //No. Casa
        Direccion_Emergencia_Separada[3], //Población
        Direccion_Emergencia_Separada[2], //Distrito
        Direccion_Emergencia_Separada[7], //Código post./ Pobl.
        Direccion_Emergencia_Separada[6], //Clave de país
        Tel_Celular_1, //Nº teléfono
        "0", //Distancia en km
        "", //Viviendas p.personal
        "", //Ruta del autobús
        "", //Cpo.adicional direc.
        "", //Calle 2
        "", //Calle 3
        Direccion_Emergencia_Separada[5], //Región
        "", //Nº (edificio)
        "", //N°Vivienda
        "", //Sigla del edificio
        "", //Piso
        "", //Sigla vía pública
        "0", //Distancia en km
        Tel_Casa != "" ? "TEL2" : Tel_Celular_2 != "" ? "CELL" : "", //Clase comunicación
        Tel_Casa != "" ? Tel_Casa : Tel_Celular_2 != "" ? Tel_Celular_2 : "", //Número
        "", //Clase comunicación
        "", //Número
        Direccion_Emergencia_Separada[4], //Código de la ciudad
        "0" //Subscrip.soc.ferrocarril
      ]);
    }

    if(Nombre_Pareja){
      Datos_Infotipo_IT0021.push([
        Numero_Colaborador, //Número de personal
        "1", //Subtipo (Información de la pareja)
        "", //Identif.objeto
        "", //Indicador bloqueo
        "31/12/9999", //Fin de validez
        "", //Inicio de validez
        "0", //Nº regist.infotipo
        "1", //Miembro de familia
        Fecha_de_Nacimiento_Pareja, //Fecha de nacimiento
        "MX", //País nacim.
        "MX", //Nacionalidad
        Sexo_Pareja === "H" ? 1 : Sexo_Pareja === "M" ? 2 : "", //Sexo
        Nombre_Pareja, //Nombre
        Apellido_Pareja, //Apellido
        "", //Lugar de nacimiento
        "", //Estado federado
        "", //Nº pers.miembro fam.
        "", //Apellido de soltera
        Segundo_Apellido_Pareja, //Segundo apellido
        Nombre_Pareja + " " + Apellido_Pareja + " " + Segundo_Apellido_Pareja, //Nombre completo
        Edad_Pareja, //Edad
        Grado_Escolar_Pareja, //Grado Escolar
        "0", //Indicador de edición
        "", //Iniciales
        "", //Prefijo
        "", //Prefijo
        Grado_Escolar_Pareja, //Título
        "", //Nº PVF
        "", //Filiación
        "", //Autoriz.
        "", //Alojamiento del hijo
        "", //DchoSubsHij
        "", //Clase hijo
        "", //Pluses por hijos
        "", //Plus por hijos
        "", //Autoriz.Certif.Enf.
        "", //Número autorización
        "0", //NASS
        "0", //NASS
        "", //Conf.hasta
        "", //Formación
        "", //Instituto
        "", //Direc.empr.
        "", //Otras nacionalidades
        "", //2ª/3ª nacionalidad
        "0", //Importe
        "", //Titulo
        "", //Persona referencia en caso de emergencia
        "", //Pluses por hijos
        "", //Plus por hijos
        "", //Autoriz.Certif.Enf.
        "", //Número autorización
        "", //NASS
        "", //NASS
        "", //Conf.hasta
        "", //Formación
        "", //Instituto
        "", //Direc.empr.
        "", //Otras nacionalidades
        "", //2ª/3ª nacionalidad
        "", //Importe
        "", //Título
        ""  //Persona referencia en caso de emergencia
      ]);
    }

    for(let j = 0; j < Familiares.length; j++){
      if(Familiares[j][0]){
        Datos_Infotipo_IT0021.push([
          Numero_Colaborador, //Número de personal
          "2", //Subtipo (Información de la pareja)
          j+1, //Identif.objeto
          "", //Indicador bloqueo
          "31/12/9999", //Fin de validez
          "", //Inicio de validez
          "0", //Nº regist.infotipo
          "2", //Miembro de familia
          Familiares[j][7], //Fecha de nacimiento
          "MX", //País nacim.
          "MX", //Nacionalidad
          Familiares[j][6] === "H" ? 1 : Familiares[j][6] === "M" ? 2 : "", //Sexo
          Familiares[j][0], //Nombre
          Familiares[j][2], //Apellido
          "", //Lugar de nacimiento
          "", //Estado federado
          "", //Nº pers.miembro fam.
          "", //Apellido de soltera
          Familiares[j][4], //Segundo apellido
          Familiares[j][0] + " " +  Familiares[j][2] + " " +  Familiares[j][4], //Nombre completo
          Familiares[j][9], //Edad
          Familiares[j][10], //Grado Escolar
          "0", //Indicador de edición
          "", //Iniciales
          "", //Prefijo
          "", //Prefijo
          Familiares[j][10], //Título
          "", //Nº PVF
          "", //Filiación
          "", //Autoriz.
          "", //Alojamiento del hijo
          "", //DchoSubsHij
          "", //Clase hijo
          "", //Pluses por hijos
          "", //Plus por hijos
          "", //Autoriz.Certif.Enf.
          "", //Número autorización
          "0", //NASS
          "0", //NASS
          "", //Conf.hasta
          "", //Formación
          "", //Instituto
          "", //Direc.empr.
          "", //Otras nacionalidades
          "", //2ª/3ª nacionalidad
          "0", //Importe
          "", //Titulo
          "", //Persona referencia en caso de emergencia
          "", //Pluses por hijos
          "", //Plus por hijos
          "", //Autoriz.Certif.Enf.
          "", //Número autorización
          "", //NASS
          "", //NASS
          "", //Conf.hasta
          "", //Formación
          "", //Instituto
          "", //Direc.empr.
          "", //Otras nacionalidades
          "", //2ª/3ª nacionalidad
          "", //Importe
          "", //Título
          ""  //Persona referencia en caso de emergencia
        ]);
      }
    }

    let Datos_Contacto_Institucional = [];
    let Datos_Uniforme_Administrativo = [];
    let Datos_Uniforme_Operativo = [];
    let Datos_Uniforme_Chamarra = [];

    Libro_Destino = SpreadsheetApp.openById(ID_Libro_Destino);

    Hoja_Destino_Contacto_Institucional = Libro_Destino.getSheetByName("Contacto Institucional");

    Hoja_Destino_Operativo = Libro_Destino.getSheetByName("Operativo");
    Hoja_Destino_Administrativo = Libro_Destino.getSheetByName("Administrativo");
    Hoja_Destino_Chamarra = Libro_Destino.getSheetByName("Chamarra");

    if (Correo_Trabajo !== "" || Celular_Trabajo !== ""){
      let Contacto_Fila = Hoja_Destino_Contacto_Institucional.getLastRow() + 1;
      Datos_Contacto_Institucional = [[Numero_Colaborador, Nombre_Colaborador, Puesto, Area, `=VLOOKUP(A${Contacto_Fila}, ZREPORTES!A:AD, 4, FALSE)`, `=VLOOKUP(A${Contacto_Fila}, ZREPORTES!A:AD, 19, FALSE)`, Fecha_de_Nacimiento, Correo_Trabajo, Celular_Trabajo]]
    }

    if(Camisa_Admon !== "" || Camisa_Admon_Manga_Larga !== "" || Pantalon_Admon !== "" || Botas_Admon !== ""){
      let Admin_Fila = Hoja_Destino_Administrativo.getLastRow() + 1;
      Datos_Uniforme_Administrativo = [[Numero_Colaborador, Nombre_Colaborador, Puesto, Area, `=VLOOKUP(A${Admin_Fila}, ZREPORTES!A:AD, 4, FALSE)`, Camisa_Admon, Camisa_Admon_Manga_Larga, "", "", Pantalon_Admon, Botas_Admon]]
    }

    if (Camisola !== "" || Pantalon_Ope !== "" || Botas !== "" || Playera_Blanca !== "") {
      let Ope_Fila = Hoja_Destino_Operativo.getLastRow() + 1;
      Datos_Uniforme_Operativo = [[Numero_Colaborador, Nombre_Colaborador, Puesto, Area, `=VLOOKUP(A${Ope_Fila}, ZREPORTES!A:AD, 4, FALSE)`, Camisola, "", Pantalon_Ope, Playera_Blanca, Botas]]
    }

    if(Chamarra_Chaleco !== ""){
      let Chamarra_Fila = Hoja_Destino_Chamarra.getLastRow() + 1;
      let Datos_Chamarra = Chamarra_Chaleco.split(",").map(parte => parte.trim());
      Datos_Uniforme_Chamarra = [[Numero_Colaborador, Nombre_Colaborador, Puesto, Area, `=VLOOKUP(A${Chamarra_Fila}, ZREPORTES!A:AD, 4, FALSE)`, Datos_Chamarra[0], Playera_Casual, Datos_Chamarra[1]]]
    }

    send_Data(Datos_Infotipo_IT0006, Datos_Infotipo_IT0021, Datos_Contacto_Institucional, Datos_Uniforme_Administrativo, Datos_Uniforme_Operativo, Datos_Uniforme_Chamarra);
  }
}

function show_Data(Pegar_Copia){
  Pegar_Copia.getRange("B6").setValue(Nombre_Colaborador);

  Pegar_Copia.getRange("B7").setValue(Area);
  Pegar_Copia.getRange("K7").setValue(Fecha_de_Nacimiento);

  Pegar_Copia.getRange("B8").setValue(Puesto);
  Pegar_Copia.getRange("K8").setValue(Telefono_Casa);
  Pegar_Copia.getRange("K9").setValue(Telefono_Celular);

  Pegar_Copia.getRange("B10").setValue(Direccion_Principal_Separada[0]);
  Pegar_Copia.getRange("D10").setValue(Direccion_Principal_Separada[1]);
  Pegar_Copia.getRange("E10").setValue(Direccion_Principal_Separada[2]);
  Pegar_Copia.getRange("G10").setValue(Direccion_Principal_Separada[3]);
  Pegar_Copia.getRange("H10").setValue(Direccion_Principal_Separada[4]);
  Pegar_Copia.getRange("I10").setValue(Direccion_Principal_Separada[5] + ", " + Direccion_Principal_Separada[6]);
  Pegar_Copia.getRange("J10").setValue(Direccion_Principal_Separada[7]);
  Pegar_Copia.getRange("L10").setValue(Ubicacion);

  Pegar_Copia.getRange("B13").setValue(Estado_Civil);

  Pegar_Copia.getRange("B16").setValue(Nombre_Pareja);
  Pegar_Copia.getRange("D16").setValue(Apellido_Pareja);
  Pegar_Copia.getRange("F16").setValue(Segundo_Apellido_Pareja);
  Pegar_Copia.getRange("H16").setValue(Sexo_Pareja);
  Pegar_Copia.getRange("I16").setValue(Fecha_de_Nacimiento_Pareja);
  Pegar_Copia.getRange("K16").setValue(Edad_Pareja);
  Pegar_Copia.getRange("L16").setValue(Grado_Escolar_Pareja);

  Pegar_Copia.getRange(17, 2, Familiares.length, Familiares[0].length).setValues(Familiares);

  Pegar_Copia.getRange("B26").setValue(Nombre_Contacto_Emergencia);
  Pegar_Copia.getRange("B27").setValue(Parentesco);

  Pegar_Copia.getRange("B28").setValue(Direccion_Emergencia_Separada[0]);
  Pegar_Copia.getRange("C28").setValue(Direccion_Emergencia_Separada[1]);
  Pegar_Copia.getRange("D28").setValue(Direccion_Emergencia_Separada[2]);
  Pegar_Copia.getRange("E28").setValue(Direccion_Emergencia_Separada[3]);
  Pegar_Copia.getRange("F28").setValue(Direccion_Emergencia_Separada[4]);
  Pegar_Copia.getRange("G28").setValue(Direccion_Emergencia_Separada[5] + ", " + Direccion_Emergencia_Separada[6]);
  Pegar_Copia.getRange("H28").setValue(Direccion_Emergencia_Separada[7]);

  Pegar_Copia.getRange("D29").setValue(Tel_Celular_1);
  Pegar_Copia.getRange("D30").setValue(Tel_Celular_2);
  Pegar_Copia.getRange("B31").setValue(Tel_Casa);

  Pegar_Copia.getRange("J26").setValue(Camisa_Admon);
  Pegar_Copia.getRange("J27").setValue(Camisa_Admon_Manga_Larga);
  Pegar_Copia.getRange("J28").setValue(Pantalon_Admon);
  Pegar_Copia.getRange("J29").setValue(Botas_Admon);
  Pegar_Copia.getRange("L26").setValue(Camisola);
  Pegar_Copia.getRange("L27").setValue(Pantalon_Ope);
  Pegar_Copia.getRange("L28").setValue(Botas);
  Pegar_Copia.getRange("L29").setValue(Playera_Blanca);
  Pegar_Copia.getRange("J30").setValue(Playera_Casual);
  Pegar_Copia.getRange("L30").setValue(Chamarra_Chaleco);

  Pegar_Copia.getRange("J31").setValue(Celular_Trabajo);
  Pegar_Copia.getRange("L31").setValue(Correo_Trabajo);
}
