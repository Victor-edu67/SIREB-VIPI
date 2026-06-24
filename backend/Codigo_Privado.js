// ID del Excel. OJO: no cambiar o se muere todo el sistema
const ID_BASE_DATOS = "1fYHYW17iMh1WLOSwVXTWTFa-AtsFDc_9TVHtAltu6RQ";

// Arranca la app web y revisa la sesión del usuario
function doGet() {
  var correoUsuario = Session.getActiveUser().getEmail();
  
  if (correoUsuario === "") {
    return HtmlService.createHtmlOutput("Por favor, inicia sesión en tu cuenta de Google para acceder al sistema SIREB-VIPI.");
  }
  
  var rol = obtenerRolUsuario(correoUsuario);
  
  // --- Parches por si la profe prueba sin configurar bien el Sheets ---
  if (rol === "Error_Falta_ID") {
    return HtmlService.createHtmlOutput("<div style='font-family: sans-serif; padding: 40px; text-align: center; color: #333;'><h2>Falta conectar el Excel</h2><p>Debes pegar el ID de tu Google Sheets en la línea 2 del archivo Codigo.gs (ID_BASE_DATOS).</p></div>");
  }
  if (rol === "Error_Falta_Pestana") {
    return HtmlService.createHtmlOutput("<div style='font-family: sans-serif; padding: 40px; text-align: center; color: #333;'><h2>Error de Conexión con Excel</h2><p>El ID es correcto, pero el sistema no encuentra la pestaña 'Usuarios'.</p></div>");
  }
  
  // Ruteo básico según el rol
  var template;
  if (rol === "Administrador") {
    template = HtmlService.createTemplateFromFile('Admin_Panel'); 
    template.correo = correoUsuario; 
  } else if (rol === "Estudiante") {
    template = HtmlService.createTemplateFromFile('Estudiante_Solvencia'); 
    template.correo = correoUsuario;
  } else {
    template = HtmlService.createTemplateFromFile('SinAcceso'); 
  }
  
  return template.evaluate()
      .setTitle('SIREB - Biblioteca VIPI')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Clava el préstamo en el Excel. Falta optimizar un poco pero aguanta la pela
function registrarPrestamoEnSheet(cedulaIngresada, idLibro, fechaSeleccionada) {
  const ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  const sheetUsuarios = ss.getSheetByName("Usuarios");
  const sheetLibros = ss.getSheetByName("Inventario_Libros"); 
  const sheetPrestamos = ss.getSheetByName("Registro_Prestamos"); 

  const datosUsuarios = sheetUsuarios.getDataRange().getValues();
  const datosLibros = sheetLibros.getDataRange().getValues();

  // 1. Validar que el estudiante exista y no deba nada (requerimiento estricto)
  let usuarioEncontrado = false;
  let estatusSancion = "";
  let filaUsuario = -1; 

  for (let i = 1; i < datosUsuarios.length; i++) {
    if (datosUsuarios[i][0] && datosUsuarios[i][0].toString().trim() === cedulaIngresada.toString().trim()) {
      usuarioEncontrado = true;
      estatusSancion = datosUsuarios[i][4]; 
      filaUsuario = i + 1; 
      break;
    }
  }

  if (!usuarioEncontrado) return { exito: false, mensaje: "Error: La cédula ingresada no pertenece a un usuario registrado." };
  if (estatusSancion === "Inhabilitado") return { exito: false, mensaje: "Denegado: El estudiante presenta una multa activa." };

  // 2. Revisar si el libro existe y no se lo llevaron ya
  let libroEncontrado = false;
  let filaLibro = -1; 
  let estatusLibro = ""; 

  for (let i = 1; i < datosLibros.length; i++) {
    if (datosLibros[i][0].toString().trim() === idLibro.toString().trim()) {
      libroEncontrado = true;
      filaLibro = i + 1; 
      estatusLibro = datosLibros[i][5] ? datosLibros[i][5].toString().trim() : ""; 
      break;
    }
  }

  if (!libroEncontrado) return { exito: false, mensaje: "Error: El código del libro no existe en el catálogo." };
  if (estatusLibro === "Prestado") return { exito: false, mensaje: "Error: Este ejemplar ya se encuentra prestado a otro usuario." };

  // 3. Fechas. Parche medio feo para saltarse sábados y domingos (costó que funcionara)
  const idPrestamo = "P-" + Utilities.getUuid().substring(0, 3).toUpperCase();
  const fechaSalida = new Date(fechaSeleccionada + "T12:00:00");
  
  let diaSalida = fechaSalida.getDay();
  if (diaSalida === 0 || diaSalida === 6) {
    return { exito: false, mensaje: "Error: No se pueden registrar préstamos con fecha de salida en sábado o domingo." };
  }

  const fechaEntregaPrevista = new Date(fechaSalida);
  let diasSumados = 0;
  
  while (diasSumados < 3) {
    fechaEntregaPrevista.setDate(fechaEntregaPrevista.getDate() + 1); 
    let diaDeLaSemana = fechaEntregaPrevista.getDay(); 
    if (diaDeLaSemana !== 0 && diaDeLaSemana !== 6) {
      diasSumados++;
    }
  }
  
  // Guardar en la BD
  sheetPrestamos.appendRow([
    idPrestamo, cedulaIngresada, idLibro, fechaSalida, fechaEntregaPrevista, "", 0, "Activo"
  ]);

  // Bloquear al estudiante temporalmente
  sheetUsuarios.getRange(filaUsuario, 5).setValue("Inhabilitado");

  // Marcar el libro como no disponible
  sheetLibros.getRange(filaLibro, 6).setValue("Prestado");

  const salidaTexto = Utilities.formatDate(fechaSalida, "GMT-4", "dd/MM/yyyy");
  const entregaTexto = Utilities.formatDate(fechaEntregaPrevista, "GMT-4", "dd/MM/yyyy");

  return { 
    exito: true, 
    mensaje: `¡Préstamo registrado con éxito!\n\nID de recibo: ${idPrestamo}\nFecha de salida: ${salidaTexto}\nDevolver el día: ${entregaTexto} (Sin contar fines de semana)` 
  };
}

// Trigger que corre de madrugada para clavar las multas (1 día = 1 semana)
function verificarRetrasosYMultas() {
  const ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  const sheetPrestamos = ss.getSheetByName("Registro_Prestamos"); 
  const sheetUsuarios = ss.getSheetByName("Usuarios");

  const datosPrestamos = sheetPrestamos.getDataRange().getValues();
  const datosUsuarios = sheetUsuarios.getDataRange().getValues();

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  // Mapeamos los usuarios para no hacer un for anidado y reventar la cuota de ejecución de Google
  let mapaUsuarios = {};
  for (let j = 1; j < datosUsuarios.length; j++) {
    let cedula = datosUsuarios[j][0].toString(); 
    mapaUsuarios[cedula] = j + 1; 
  }

  for (let i = 1; i < datosPrestamos.length; i++) {
    let estadoPrestamo = datosPrestamos[i][7]; // Índice 7 basado en el appendRow
    if (estadoPrestamo === "Activo") {
      let fechaEntregaPrevista = new Date(datosPrestamos[i][4]); 
      fechaEntregaPrevista.setHours(0, 0, 0, 0);

      if (fechaEntregaPrevista < hoy) {
        let diferenciaMilisegundos = hoy.getTime() - fechaEntregaPrevista.getTime();
        let diasRetraso = Math.floor(diferenciaMilisegundos / (1000 * 60 * 60 * 24));
        
        if (diasRetraso > 0) {
          let cedulaUsuario = datosPrestamos[i][1].toString(); 
          let filaUsuario = mapaUsuarios[cedulaUsuario];
          
          if (filaUsuario) {
            let semanasMulta = diasRetraso; 
            sheetUsuarios.getRange(filaUsuario, 5).setValue("Inhabilitado");
            sheetUsuarios.getRange(filaUsuario, 6).setValue(semanasMulta);
          }
        }
      }
    }
  }
}

// Agarra la plantilla de Docs y le clava los datos. Tarda un poco.
function generarSolvenciaPDF(correoEstudiante) {
  var idPlantilla = '1JXxS5wnU9S7cl1xeEjXW9DXeqSzk8AYUKzZosoN4-e4'; 
  var datosEstudiante = obtenerDatosPorCorreo(correoEstudiante); 
  
  if (!datosEstudiante) {
    return { exito: false, mensaje: "Error: No se encontraron los datos de este correo en la base de datos. Verifica la pestaña Usuarios." };
  }
  
  if (datosEstudiante.estatusSancion !== "Solvente") {
    return { exito: false, mensaje: "No se puede emitir la solvencia. Presenta multas activas." };
  }
  
  // TODO: Buscar forma de no crear un archivo temporal porque llena el Drive rápido
  var archivoPlantilla = DriveApp.getFileById(idPlantilla);
  var copiaDoc = archivoPlantilla.makeCopy("Temp_Solvencia");
  var idCopia = copiaDoc.getId();
  
  var doc = DocumentApp.openById(idCopia);
  var cuerpo = doc.getBody();
  var fechaActual = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy");
  
  cuerpo.replaceText("{{NOMBRE_ESTUDIANTE}}", datosEstudiante.nombre);
  cuerpo.replaceText("{{CEDULA}}", datosEstudiante.cedula);
  cuerpo.replaceText("{{CARRERA}}", datosEstudiante.carrera);
  cuerpo.replaceText("{{FECHA}}", fechaActual);
  
  doc.saveAndClose(); 
  
  var pdfBlob = copiaDoc.getAs('application/pdf');
  copiaDoc.setTrashed(true); // Borrar la evidencia jaja
  
  var pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  var nombreFinal = "Solvencia_Biblioteca_" + datosEstudiante.cedula + ".pdf";
  
  return { exito: true, archivoBase64: pdfBase64, nombreArchivo: nombreFinal };
}

// Utils para buscar info rápido
function obtenerRolUsuario(correoBuscado) {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS); 
  var sheet = ss.getSheetByName("Usuarios");
  var datos = sheet.getDataRange().getValues();
  
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][6].toString().toLowerCase() === correoBuscado.toLowerCase()) {
      return datos[i][7]; 
    }
  }
  return "No Registrado";
}

function obtenerDatosPorCorreo(correoBuscado) {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS); 
  var sheet = ss.getSheetByName("Usuarios");
  var datos = sheet.getDataRange().getValues();
  
  for (var i = 1; i < datos.length; i++) {
    if (datos[i][6].toString().toLowerCase() === correoBuscado.toLowerCase()) {
      return {
        cedula: datos[i][0].toString(),
        nombre: datos[i][1].toString(),
        carrera: datos[i][2].toString(),
        estatusSancion: datos[i][4].toString(),
        correo: datos[i][6].toString(),
        rol: datos[i][7].toString()
      };
    }
  }
  return null;
}

// =======================================================================
// ENDPOINTS Y QUERIES MEDIOS FEOS PERO FUNCIONALES (NO TOCAR)
// =======================================================================

function obtenerDatosYPrestamos(correoBuscado) {
  var datosBasicos = obtenerDatosPorCorreo(correoBuscado);
  if (!datosBasicos) return null;

  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheetPrestamos = ss.getSheetByName("Registro_Prestamos");
  var prestamosData = sheetPrestamos.getDataRange().getValues();

  datosBasicos.prestamosActivos = [];

  for (var j = 1; j < prestamosData.length; j++) {
    if (prestamosData[j][1].toString() === datosBasicos.cedula && prestamosData[j][7] === "Activo") {
      var fVenc = new Date(prestamosData[j][4]);
      datosBasicos.prestamosActivos.push({
        idPrestamo: prestamosData[j][0],
        idLibro: prestamosData[j][2],
        fechaVencimiento: Utilities.formatDate(fVenc, "GMT-4", "dd/MM/yyyy")
      });
    }
  }
  return datosBasicos;
}

function buscarPrestamosParaDevolucion(cedula) {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheetPrestamos = ss.getSheetByName("Registro_Prestamos");
  var prestamosData = sheetPrestamos.getDataRange().getValues();
  
  var activos = [];
  
  for (var j = 1; j < prestamosData.length; j++) {
    // Índice 1: Cédula | Índice 7: Estatus Activo
    if (prestamosData[j][1].toString() === cedula.toString() && prestamosData[j][7] === "Activo") {
      var fSalida = new Date(prestamosData[j][3]); // Índice 3: Fecha de Salida
      var fVenc = new Date(prestamosData[j][4]);   // Índice 4: Fecha Prevista
      
      activos.push({
        idPrestamo: prestamosData[j][0],
        idLibro: prestamosData[j][2],
        fechaVencimiento: Utilities.formatDate(fVenc, "GMT-4", "dd/MM/yyyy"),
        // Lo mandamos en formato gringo para que el HTML lo agarre bien
        fechaSalida: Utilities.formatDate(fSalida, "GMT-4", "yyyy-MM-dd") 
      });
    }
  }
  
  if(activos.length === 0) return { exito: false, mensaje: "El estudiante no tiene préstamos activos." };
  return { exito: true, prestamos: activos };
}

function obtenerTodosLosPrestamosActivos() {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheetPrestamos = ss.getSheetByName("Registro_Prestamos");
  var dataP = sheetPrestamos.getDataRange().getValues();

  var activos = [];
  for (var i = 1; i < dataP.length; i++) {
    if (dataP[i][7] === "Activo") { 
      var fSalida = new Date(dataP[i][3]); 
      var fVenc = new Date(dataP[i][4]);  
      
      activos.push({
        idPrestamo: dataP[i][0],
        cedula: dataP[i][1],
        idLibro: dataP[i][2],
        fechaVencimiento: Utilities.formatDate(fVenc, "GMT-4", "dd/MM/yyyy"),
        fechaSalida: Utilities.formatDate(fSalida, "GMT-4", "yyyy-MM-dd")
      });
    }
  }
  return activos;
}

// Mata el préstamo en la BD
function procesarDevolucionEnSheet(idPrestamo, fechaDevolucion) { 
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheetPrestamos = ss.getSheetByName("Registro_Prestamos");
  var sheetUsuarios = ss.getSheetByName("Usuarios");
  var sheetLibros = ss.getSheetByName("Inventario_Libros");
  
  var dataP = sheetPrestamos.getDataRange().getValues();
  var dataU = sheetUsuarios.getDataRange().getValues();
  var dataL = sheetLibros.getDataRange().getValues();
  
  // Convertir el string del HTML a Date. Le sumamos 12 horas por un bug raro con la zona horaria.
  var fechaRealDevolucion = new Date(fechaDevolucion + "T12:00:00");
  
  for (var i = 1; i < dataP.length; i++) {
    if (dataP[i][0] === idPrestamo && dataP[i][7] === "Activo") {
      var filaPrestamo = i + 1;
      var cedula = dataP[i][1];
      var idLibro = dataP[i][2];
      
      // 1. Matar el préstamo
      sheetPrestamos.getRange(filaPrestamo, 6).setValue(fechaRealDevolucion); 
      sheetPrestamos.getRange(filaPrestamo, 8).setValue("Devuelto"); 
      
      // 2. Limpiar al Estudiante
      for(var j = 1; j < dataU.length; j++){
        if(dataU[j][0].toString() === cedula.toString()){
          sheetUsuarios.getRange(j + 1, 5).setValue("Solvente");
          sheetUsuarios.getRange(j + 1, 6).clearContent(); 
          break;
        }
      }
      
      // 3. Devolver el Libro
      for(var k = 1; k < dataL.length; k++){
        if(dataL[k][0].toString() === idLibro.toString()){
          sheetLibros.getRange(k + 1, 6).setValue("Disponible");
          break;
        }
      }
      
      return { exito: true, mensaje: "Devolución registrada exitosamente con fecha: " + Utilities.formatDate(fechaRealDevolucion, "GMT-4", "dd/MM/yyyy") + ". El usuario y el libro han sido liberados." };
    }
  }
  return { exito: false, mensaje: "No se encontró el préstamo activo." };
}

// Filtra los datos para la gráfica del admin
function obtenerEstadisticasRango(fechaInicio, fechaFin) {
  var sheetPrestamos = SpreadsheetApp.openById(ID_BASE_DATOS).getSheetByName("Registro_Prestamos");
  var data = sheetPrestamos.getDataRange().getValues();
  
  var fInicio = new Date(fechaInicio + "T00:00:00");
  var fFin = new Date(fechaFin + "T23:59:59");
  
  var stats = { total: 0, activos: 0, devueltos: 0 };
  
  for (var i = 1; i < data.length; i++) {
    if(data[i][3]) {
      var fPrestamo = new Date(data[i][3]);
      if (fPrestamo >= fInicio && fPrestamo <= fFin) {
        stats.total++;
        if (data[i][7] === "Activo") stats.activos++;
        if (data[i][7] === "Devuelto") stats.devueltos++;
      }
    }
  }
  return stats;
}