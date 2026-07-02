function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Menú de acciones')
      .addItem('Mostrar datos', 'menuItem1')
      .addItem('Guardar datos', 'menuItem2')
      .addToUi();
}

function menuItem1() {
  const Usuario = Session.getActiveUser().getEmail();
  console.log("Mostrar datos");

  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     '¿Estás seguro de esta acción?',
     'Al continuar se borrarán los datos del formato y se cargarán los nuevos',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmación recibida, cargando los datos...');
    try{
      get_Data(Usuario, "1");
      ui.alert('Los datos fueron cargados correctamente.');
    }catch(e){
      ui.alert(`Ocurrió un error durante la carga.\nDetalle: ${e.message}`);
    }

  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Se canceló la acción.');  
  }
}

function menuItem2() {
  const Usuario = Session.getActiveUser().getEmail();
  console.log("Guardar datos");
  
  var ui = SpreadsheetApp.getUi(); // Same variations.

  var result = ui.alert(
     '¿Estás seguro de esta acción?',
     'Al continuar se actualizarán los infotipos',
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmación recibida, guardando los datos...');
    try{
      get_Data(Usuario, "2");
    }catch(e){
      ui.alert(`Ocurrió un error durante el guardado.\nDetalle: ${e.message}`);
    }

  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Se canceló la acción.');  
  }
}
