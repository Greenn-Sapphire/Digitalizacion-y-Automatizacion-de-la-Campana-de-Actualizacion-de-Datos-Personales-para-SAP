const ID_Libro_Destino = "1Up_MnEBaEYvdTLLNuJp2dRZSIheKq0jfh40O0BG1OcU"

function send_Data(Datos_Info_06, Datos_Info_021, Datos_Contacto, Datos_Admin, Datos_Ope, Datos_Chamarra) {
  const Libro_Destino = SpreadsheetApp.openById(ID_Libro_Destino);

  const Hoja_Destino_Infotipo_IT0006 = Libro_Destino.getSheetByName("Infotipo IT0006");
  const Hoja_Destino_Infotipo_IT0021 = Libro_Destino.getSheetByName("Infotipo IT0021");

  const Hoja_Destino_Contacto_Institucional = Libro_Destino.getSheetByName("Contacto Institucional");

  const Hoja_Destino_Operativo = Libro_Destino.getSheetByName("Operativo");
  const Hoja_Destino_Administrativo = Libro_Destino.getSheetByName("Administrativo");
  const Hoja_Destino_Chamarra = Libro_Destino.getSheetByName("Chamarra");

  if (Datos_Info_06.length > 0) {
    Hoja_Destino_Infotipo_IT0006.getRange(Hoja_Destino_Infotipo_IT0006.getLastRow() + 1, 1, Datos_Info_06.length, Datos_Info_06[0].length).setValues(Datos_Info_06);
  }

  if (Datos_Info_021.length > 0) {
    Hoja_Destino_Infotipo_IT0021.getRange(Hoja_Destino_Infotipo_IT0021.getLastRow() + 1, 1, Datos_Info_021.length, Datos_Info_021[0].length).setValues(Datos_Info_021);
  }

  if (Datos_Contacto.length > 0) {
    Hoja_Destino_Contacto_Institucional.getRange(Hoja_Destino_Contacto_Institucional.getLastRow() + 1, 1, Datos_Contacto.length, Datos_Contacto[0].length).setValues(Datos_Contacto);
  }

  if (Datos_Admin.length > 0) {
    Hoja_Destino_Administrativo.getRange(Hoja_Destino_Administrativo.getLastRow() + 1, 1, Datos_Admin.length, Datos_Admin[0].length).setValues(Datos_Admin);
  }

  if (Datos_Ope.length > 0) {
    Hoja_Destino_Operativo.getRange(Hoja_Destino_Operativo.getLastRow() + 1, 1, Datos_Ope.length, Datos_Ope[0].length).setValues(Datos_Ope);
  }

  if (Datos_Chamarra.length > 0) {
    Hoja_Destino_Chamarra.getRange(Hoja_Destino_Chamarra.getLastRow() + 1, 1, Datos_Chamarra.length, Datos_Chamarra[0].length).setValues(Datos_Chamarra);
  }
}
