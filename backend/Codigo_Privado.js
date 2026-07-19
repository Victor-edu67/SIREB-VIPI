// ID del Excel. OJO: no cambiar o se muere todo el sistema
const ID_BASE_DATOS = "1fYHYW17iMh1WLOSwVXTWTFa-AtsFDc_9TVHtAltu6RQ";

// Arranca la app web y revisa la sesión del usuario
function doGet() {
  var correoUsuario = Session.getActiveUser().getEmail();
  
  if (correoUsuario === "") {
    return HtmlService.createHtmlOutput("Por favor, inicia sesión en tu cuenta de Google para acceder al sistema SIREB-VIPI.");
  }
  
  var rol = obtenerRolUsuario(correoUsuario);
  
  if (rol === "Error_Falta_ID") {
    return HtmlService.createHtmlOutput("<div style='font-family: sans-serif; padding: 40px; text-align: center; color: #333;'><h2>Falta conectar el Excel</h2><p>Debes pegar el ID de tu Google Sheets en la línea 2 del archivo Codigo.gs (ID_BASE_DATOS).</p></div>");
  }
  if (rol === "Error_Falta_Pestana") {
    return HtmlService.createHtmlOutput("<div style='font-family: sans-serif; padding: 40px; text-align: center; color: #333;'><h2>Error de Conexión con Excel</h2><p>El ID es correcto, pero el sistema no encuentra la pestaña 'Usuarios'.</p></div>");
  }
  
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

// Registra el préstamo mapeado a las 8 columnas exactas de Registro_Prestamos
function registrarPrestamoEnSheet(cedulaIngresada, idLibro, fechaSeleccionada) {
  const ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  const sheetUsuarios = ss.getSheetByName("Usuarios");
  const sheetLibros = ss.getSheetByName("Inventario_Libros"); 
  const sheetPrestamos = ss.getSheetByName("Registro_Prestamos"); 

  const datosUsuarios = sheetUsuarios.getDataRange().getValues();
  const datosLibros = sheetLibros.getDataRange().getValues();

  let usuarioEncontrado = false;
  let estatusSancion = "";
  let filaUsuario = -1; 

  for (let i = 1; i < datosUsuarios.length; i++) {
    if (datosUsuarios[i][0] && datosUsuarios[i][0].toString().trim() === cedulaIngresada.toString().trim()) {
      usuarioEncontrado = true;
      estatusSancion = datosUsuarios[i][4]; // Estatus_Solvencia (Columna E)
      filaUsuario = i + 1; 
      break;
    }
  }

  if (!usuarioEncontrado) return { exito: false, mensaje: "Error: La cédula ingresada no pertenece a un usuario registrado." };
  if (estatusSancion === "Inhabilitado") return { exito: false, mensaje: "Denegado: El estudiante presenta una multa activa." };

  let libroEncontrado = false;
  let filaLibro = -1; 
  let estatusLibro = ""; 

  for (let i = 1; i < datosLibros.length; i++) {
    if (datosLibros[i][0].toString().trim() === idLibro.toString().trim()) {
      libroEncontrado = true;
      filaLibro = i + 1; 
      estatusLibro = datosLibros[i][5] ? datosLibros[i][5].toString().trim() : ""; // Estado (Columna F)
      break;
    }
  }

  if (!libroEncontrado) return { exito: false, mensaje: "Error: El código del libro no existe en el catálogo." };
  if (estatusLibro === "Prestado") return { exito: false, mensaje: "Error: Este ejemplar ya se encuentra prestado a otro usuario." };

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
  
  // Guardar en Registro_Prestamos (8 Columnas exactas)
  // ID_Prestamo, Cedula_Usuario, ID_Libro, Fecha_Prestamo, Fecha_Vencimiento, Fecha_Devolucion, # Dias_Retraso, Estado_Prestamo
  sheetPrestamos.appendRow([
    idPrestamo, cedulaIngresada, idLibro, fechaSalida, fechaEntregaPrevista, "", 0, "Activo"
  ]);

  sheetUsuarios.getRange(filaUsuario, 5).setValue("Inhabilitado"); // Col E
  sheetLibros.getRange(filaLibro, 6).setValue("Prestado"); // Col F

  const salidaTexto = Utilities.formatDate(fechaSalida, "GMT-4", "dd/MM/yyyy");
  const entregaTexto = Utilities.formatDate(fechaEntregaPrevista, "GMT-4", "dd/MM/yyyy");

  return { 
    exito: true, 
    mensaje: `¡Préstamo registrado con éxito!\n\nID de recibo: ${idPrestamo}\nFecha de salida: ${salidaTexto}\nDevolver el día: ${entregaTexto} (Sin contar fines de semana)` 
  };
}

// Mata el préstamo en la BD, calcula la multa entera y respeta la estructura del Excel
function procesarDevolucionEnSheet(idPrestamo, fechaDevolucion) { 
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheetPrestamos = ss.getSheetByName("Registro_Prestamos");
  var sheetUsuarios = ss.getSheetByName("Usuarios");
  var sheetLibros = ss.getSheetByName("Inventario_Libros");
  
  var dataP = sheetPrestamos.getDataRange().getValues();
  var dataU = sheetUsuarios.getDataRange().getValues();
  var dataL = sheetLibros.getDataRange().getValues();
  
  var fechaRealDevolucion = new Date(fechaDevolucion + "T12:00:00");
  
  for (var i = 1; i < dataP.length; i++) {
    if (dataP[i][0] === idPrestamo && dataP[i][7] === "Activo") {
      var filaPrestamo = i + 1;
      var cedula = dataP[i][1];
      var idLibro = dataP[i][2];
      
      var fVenc = new Date(dataP[i][4]);
      fVenc.setHours(12, 0, 0, 0);
      
      var diasRetraso = 0;
      var estadoFinalPrestamo = "Devuelto";
      var estatusNuevo = "Solvente";
      var semanasMulta = 0; // Columna F en Usuarios espera un número entero
      var mensajeAdicional = " El usuario y el libro han sido liberados.";

      if (fechaRealDevolucion > fVenc) {
        var difMilisegundos = Math.abs(fechaRealDevolucion - fVenc);
        diasRetraso = Math.ceil(difMilisegundos / (1000 * 60 * 60 * 24));
        
        if(diasRetraso > 0) {
          estadoFinalPrestamo = "Devuelto con Multa";
          estatusNuevo = "Inhabilitado";
          semanasMulta = diasRetraso; // 1 día de retraso = 1 semana de multa (número entero)
          mensajeAdicional = `\n\n⚠️ ATENCIÓN: Devolución con ${diasRetraso} día(s) de retraso.\nSe aplicó sanción de ${semanasMulta} semanas.`;
        }
      }

      // 1. Actualizar Préstamo (Columnas F, G, H)
      sheetPrestamos.getRange(filaPrestamo, 6).setValue(fechaRealDevolucion); // Col F: Fecha Devolución
      sheetPrestamos.getRange(filaPrestamo, 7).setValue(diasRetraso);         // Col G: # Dias_Retraso
      sheetPrestamos.getRange(filaPrestamo, 8).setValue(estadoFinalPrestamo); // Col H: Estado_Prestamo
      
      // 2. Limpiar o Castigar al Estudiante (Columnas E y F)
      for(var j = 1; j < dataU.length; j++){
        if(dataU[j][0].toString() === cedula.toString()){
          sheetUsuarios.getRange(j + 1, 5).setValue(estatusNuevo);
          sheetUsuarios.getRange(j + 1, 6).setValue(semanasMulta); // Col F: Insertamos el número 0, 1, 3, etc.
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
      
      return { exito: true, mensaje: "Devolución procesada: " + Utilities.formatDate(fechaRealDevolucion, "GMT-4", "dd/MM/yyyy") + "." + mensajeAdicional };
    }
  }
  return { exito: false, mensaje: "No se encontró el préstamo activo." };
}

// Registra al estudiante respetando las columnas exactas (Incluso Carnet Vigente)
function agregarEstudianteDB(cedula, nombre, correo) {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var hojaUsuarios = ss.getSheetByName("Usuarios");
  
  // Columnas: Cedula | Nombre | Carrera | Carnet_Vigente | Estatus | Semanas_Multa | Correo | Rol
  hojaUsuarios.appendRow([cedula, nombre, "", "SI", "Solvente", 0, correo, "Estudiante"]);
  
  return "¡Estudiante registrado exitosamente en la base de datos!";
}

// Trigger que corre de madrugada para clavar las multas y marcar "Vencido"
function verificarRetrasosYMultas() {
  const ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  const sheetPrestamos = ss.getSheetByName("Registro_Prestamos"); 
  const sheetUsuarios = ss.getSheetByName("Usuarios");

  const datosPrestamos = sheetPrestamos.getDataRange().getValues();
  const datosUsuarios = sheetUsuarios.getDataRange().getValues();

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  let mapaUsuarios = {};
  for (let j = 1; j < datosUsuarios.length; j++) {
    let cedula = datosUsuarios[j][0].toString(); 
    mapaUsuarios[cedula] = j + 1; 
  }

  for (let i = 1; i < datosPrestamos.length; i++) {
    let estadoPrestamo = datosPrestamos[i][7]; 
    if (estadoPrestamo === "Activo") {
      let fechaEntregaPrevista = new Date(datosPrestamos[i][4]); 
      fechaEntregaPrevista.setHours(0, 0, 0, 0);

      if (fechaEntregaPrevista < hoy) {
        let diferenciaMilisegundos = hoy.getTime() - fechaEntregaPrevista.getTime();
        let diasRetraso = Math.floor(diferenciaMilisegundos / (1000 * 60 * 60 * 24));
        
        if (diasRetraso > 0) {
          // Marcamos el préstamo como "Vencido" y calculamos días de retraso actuales
          sheetPrestamos.getRange(i + 1, 7).setValue(diasRetraso);
          sheetPrestamos.getRange(i + 1, 8).setValue("Vencido");

          let cedulaUsuario = datosPrestamos[i][1].toString(); 
          let filaUsuario = mapaUsuarios[cedulaUsuario];
          
          if (filaUsuario) {
            let semanasMulta = diasRetraso; 
            sheetUsuarios.getRange(filaUsuario, 5).setValue("Inhabilitado");
            sheetUsuarios.getRange(filaUsuario, 6).setValue(semanasMulta); // Insertamos el número entero
          }
        }
      }
    }
  }
}

// Generador de PDF
function generarSolvenciaPDF(correoEstudiante) {
  var idPlantilla = '1JXxS5wnU9S7cl1xeEjXW9DXeqSzk8AYUKzZosoN4-e4'; 
  var datosEstudiante = obtenerDatosPorCorreo(correoEstudiante); 
  
  if (!datosEstudiante) {
    return { exito: false, mensaje: "Error: No se encontraron los datos de este correo en la base de datos." };
  }
  
  if (datosEstudiante.estatusSancion !== "Solvente") {
    return { exito: false, mensaje: "No se puede emitir la solvencia. Presenta multas activas." };
  }
  
  var archivoPlantilla = DriveApp.getFileById(idPlantilla);
  var copiaDoc = archivoPlantilla.makeCopy("Temp_Solvencia");
  var idCopia = copiaDoc.getId();
  
  var doc = DocumentApp.openById(idCopia);
  var cuerpo = doc.getBody();
  var fechaActual = Utilities.formatDate(new Date(), "GMT-4", "dd/MM/yyyy");
  
  cuerpo.replaceText("{{NOMBRE_ESTUDIANTE}}", datosEstudiante.nombre);
  cuerpo.replaceText("{{CEDULA}}", datosEstudiante.cedula);
  cuerpo.replaceText("{{FECHA}}", fechaActual);

  var textoCarrera = datosEstudiante.carrera.trim();
  if (textoCarrera === "") {
    cuerpo.replaceText(" de la carrera \\{\\{CARRERA\\}\\}", "");
    cuerpo.replaceText("\\{\\{CARRERA\\}\\}", ""); 
  } else {
    cuerpo.replaceText("\\{\\{CARRERA\\}\\}", textoCarrera);
  }
  
  doc.saveAndClose(); 
  
  var pdfBlob = copiaDoc.getAs('application/pdf');
  copiaDoc.setTrashed(true);
  
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

// ENDPOINTS
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
    if (prestamosData[j][1].toString() === cedula.toString() && prestamosData[j][7] === "Activo") {
      var fSalida = new Date(prestamosData[j][3]); 
      var fVenc = new Date(prestamosData[j][4]);   
      
      activos.push({
        idPrestamo: prestamosData[j][0],
        idLibro: prestamosData[j][2],
        fechaVencimiento: Utilities.formatDate(fVenc, "GMT-4", "dd/MM/yyyy"),
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
        // Cuenta tanto los devueltos normales como los devueltos con multa
        if (data[i][7].toString().includes("Devuelto")) stats.devueltos++;
      }
    }
  }
  return stats;
}