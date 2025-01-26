function validarCoordenadas() {
  // Obtenemos la hoja activa de Google Sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  
  // Obtenemos el rango de datos de la hoja
  var dataRangeAll = sheet.getDataRange();
  
  // Obtenemos la última fila con datos en la hoja
  var ultimaFila = dataRangeAll.getLastRow();

  // Recorremos todas las filas a partir de la segunda
  for (var i = 2; i <= ultimaFila; i++) {
    // Obtenemos el nombre de la ciudad y las coordenadas desde las columnas correspondientes
    var ciudad = sheet.getRange(i, 25).getValue(); // Columna Y (índice 25)
    var coordenadas = sheet.getRange(i, 26).getValue(); // Columna Z (índice 26)

    // Verificamos si las coordenadas están vacías
    if (!coordenadas) {
      Logger.log("Coordenadas no válidas en la fila " + i); // Si están vacías, las marcamos como no válidas
      continue;
    }

    // Dividimos las coordenadas en latitud y longitud separadas por coma
    var partes = coordenadas.split(",");
    var latitud = parseFloat(partes[0].trim()); // Obtenemos la latitud
    var longitud = parseFloat(partes[1].trim()); // Obtenemos la longitud

    // Verificamos si las coordenadas no son números válidos
    if (isNaN(latitud) || isNaN(longitud)) {
      Logger.log("Coordenadas no válidas en la fila " + i); // Si no son números válidos, las marcamos como no válidas
      continue;
    }

    // Redondeamos las coordenadas a 6 decimales para precisión
    latitud = parseFloat(latitud.toFixed(6));
    longitud = parseFloat(longitud.toFixed(6));

    // Utilizamos la geocodificación inversa para obtener la dirección de las coordenadas
    var geocoder = Maps.newGeocoder().reverseGeocode(latitud, longitud);
    var resultado = geocoder.results[0]; // Obtenemos el primer resultado de la geocodificación inversa

    // Si se obtiene un resultado de geocodificación
    if (resultado) {
      var direccionObtenida = resultado.formatted_address; // Obtenemos la dirección formateada

      // Verificamos si la ciudad obtenida de la dirección coincide con la ciudad proporcionada
      if (direccionObtenida.includes(ciudad)) {
        Logger.log("Coordenadas válidas para la fila " + i); // Si la ciudad coincide, marcamos las coordenadas como válidas
      } else {
        Logger.log("Coordenadas no válidas para la fila " + i + ". Dirección obtenida: " + direccionObtenida);
        
        // Si la ciudad no coincide, marcamos la celda de coordenadas en rojo
        sheet.getRange(i, 26).setBackground("red");

        // Añadimos un comentario en la celda con la dirección obtenida
        sheet.getRange(i, 26).setComment("Dirección obtenida: " + direccionObtenida);
      }
    } else {
      // Si no se pudo obtener la dirección para las coordenadas, lo registramos en el log
      Logger.log("No se pudo obtener la dirección para las coordenadas en la fila " + i);
    }

    // Pausamos 1 segundo entre cada llamada a la API de geocodificación para evitar sobrecargar el servicio
    Utilities.sleep(1000);
  }
}
