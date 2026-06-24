const ID_BASE_DATOS = "1fYHYW17iMh1WLOSwVXTWTFa-AtsFDc_9TVHtAltu6RQ";

// Carga la vista del catálogo (ALLOWALL para que funcione en Sites)
function doGet() {
  return HtmlService.createTemplateFromFile('CatalogoPublico')
      .evaluate()
      .setTitle('Catálogo Biblioteca VIPI')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Buscador principal (filtra por titulo, autor o cota LCC)
function buscarLibrosEnSheet(query) {
  var ss = SpreadsheetApp.openById(ID_BASE_DATOS);
  var sheet = ss.getSheetByName("Inventario_Libros");
  var datos = sheet.getDataRange().getValues();
  var resultados = [];
  
  var q = query.toString().toLowerCase();
  
  // i=1 para saltar los encabezados del excel
  for (var i = 1; i < datos.length; i++) {
    
    // Ojo con los indices de las columnas si modifican el archivo
    var lcc = datos[i][3] ? datos[i][3].toString().toLowerCase() : ""; 
    var titulo = datos[i][1] ? datos[i][1].toString().toLowerCase() : "";
    var autor = datos[i][2] ? datos[i][2].toString().toLowerCase() : "";
    var tipo = datos[i][3] ? datos[i][3] : ""; 
    
    var estado = datos[i][5] ? datos[i][5].toString() : "Consultar";
    
    if (titulo.includes(q) || autor.includes(q) || lcc.includes(q)) {
      resultados.push({
        titulo: datos[i][1], 
        autor: datos[i][2], 
        lcc: datos[i][3],
        tipo: tipo, 
        estado: estado
      });
    }
  }
  return resultados;
}