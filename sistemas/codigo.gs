// ==================== CONFIGURACIÓN ====================
var NOMBRE_EMPRESA = "mi_eelo";
var NIT_EMPRESA = "NIT: 12345678-9";
var DIRECCION = "Creamos Guatemala";
var HORAS_JORNADA_NORMAL = 7; // Horas de trabajo por día

// URL de tu exportación de KoboToolbox
var URL_KOBO = "https://kf.kobotoolbox.org/api/v2/assets/agi395bJj6ojXJzPPDT9n6/export-settings/es4oUjEmPvovgLd6Y5yrQ4K/data.csv";

// Nombres de los días en español (0=Domingo, 1=Lunes, ..., 6=Sábado)
var DIAS_SEMANA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

// ==================== MENÚ Y TRIGGERS ====================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⏱️ Control de Horas')
    .addItem('🚀 INSTALAR TODO', 'instalarTodo')
    .addSeparator()
    .addItem('🔄 Actualizar desde Kobo', 'importarCSVdesdeKobo')
    .addSeparator()
    .addItem('📅 Reporte por Día', 'generarReportePorDia')
    .addItem('📆 Reporte por Semana', 'generarReportePorSemana')
    .addItem('🗓️ Reporte por Mes', 'generarReportePorMes')
    .addItem('📊 Reporte por Rango', 'generarReportePorRango')
    .addItem('📋 Reporte Completo', 'generarReporteTodo')
    .addSeparator()
    .addItem('✨ NORMALIZAR TODO', 'normalizarTodo')
    .addSeparator()
    .addItem('🔀 Normalizar Nombres', 'crearHojaNombresCanonicos')
    .addItem('✏️ Cambiar Nombre de Participante', 'cambiarNombreParticipante')
    .addItem('📚 Configurar Días de Estudio', 'crearHojaDiasEstudio')
    .addItem('🧘 Configurar Lista de Terapias', 'crearHojaListaTerapias')
    .addItem('⚙️ Configurar actualización automática', 'configurarActualizacionAutomatica')
    .addSeparator()
    .addItem('🔧 Reparar Datos Kobo', 'repararDatosKobo')
    .addItem('🔍 Diagnosticar Datos Kobo', 'diagnosticarDatosKobo')
    .addToUi();
}

// ==================== NORMALIZACIÓN AUTOMÁTICA DE ACCIÓN ====================
// Normaliza silenciosamente la columna de acción en DatosKobo:
// → 🟢 Entrada (verde) / 🔴 Salida (rojo)
// Se llama automáticamente al importar datos desde Kobo.
function normalizarAccionSilencioso(hoja) {
  try {
    var datos = hoja.getDataRange().getValues();
    if (datos.length < 2) return;
    var cols = detectarColumnas(datos[0], datos.slice(1));
    if (cols.accionUnificada === undefined) return;

    // Asegurar columna subtipo_egreso
    var colSubtipo = cols.subtipoEgreso;
    if (colSubtipo === undefined) {
      var numCols = datos[0].length;
      hoja.getRange(1, numCols + 1).setValue('subtipo_egreso')
          .setFontWeight('bold').setBackground('#e6b8a2').setFontColor('#000000');
      colSubtipo = numCols;
      cols.subtipoEgreso = colSubtipo;
    }

    var SUBTIPOS_MAP = [
      { clave: 'terapia', subtipo: 'Terapia' },
      { clave: 'permiso', subtipo: 'Permiso' },
      { clave: 'comput',  subtipo: 'Computacion' }
    ];

    for (var f = 1; f < datos.length; f++) {
      var valRaw = String(datos[f][cols.accionUnificada] || '').trim();
      if (!valRaw) continue;

      var tipo = obtenerTipoRegistro(datos[f], cols);
      var correcto = null;
      var subtipo = '';

      // Detectar subtipo desde valor original
      var valLow = valRaw.toLowerCase();
      for (var s = 0; s < SUBTIPOS_MAP.length; s++) {
        if (valLow.indexOf(SUBTIPOS_MAP[s].clave) !== -1) {
          subtipo = SUBTIPOS_MAP[s].subtipo;
          break;
        }
      }

      if (tipo.esIngreso) {
        correcto = '🟢 Entrada';
        subtipo = '';
      } else if (tipo.esEgreso) {
        correcto = '🔴 Salida';
      }

      if (correcto) {
        var celda = hoja.getRange(f + 1, cols.accionUnificada + 1);
        if (valRaw !== correcto) celda.setValue(correcto);
        if (correcto === '🟢 Entrada') {
          celda.setBackground('#b7e1cd').setFontColor('#0b5c30').setFontWeight('bold');
        } else {
          celda.setBackground('#f4cccc').setFontColor('#7f0000').setFontWeight('bold');
        }
      }

      // Guardar subtipo si la celda está vacía
      if (subtipo && !String(datos[f][colSubtipo] || '').trim()) {
        hoja.getRange(f + 1, colSubtipo + 1).setValue(subtipo);
      }
    }

    // Formatear columnas start y end: convertir ISO string → fecha real legible
    var FMT_FECHA = 'dd/MM/yyyy HH:mm';
    var colsParaFecha = [];
    if (cols.start !== undefined) colsParaFecha.push(cols.start);
    if (cols.end   !== undefined) colsParaFecha.push(cols.end);

    for (var c = 0; c < colsParaFecha.length; c++) {
      var colIdx = colsParaFecha[c];
      // Aplicar formato a toda la columna (filas de datos)
      hoja.getRange(2, colIdx + 1, datos.length - 1, 1).setNumberFormat(FMT_FECHA);

      // Convertir celdas que aún sean strings ISO a objetos Date reales
      for (var f = 1; f < datos.length; f++) {
        var val = datos[f][colIdx];
        if (typeof val === 'string' && val.indexOf('T') !== -1) {
          var fechaReal = new Date(val);
          if (!isNaN(fechaReal.getTime())) {
            hoja.getRange(f + 1, colIdx + 1).setValue(fechaReal);
          }
        }
      }
    }

  } catch(e) {
    // Silencioso: no interrumpir la importación por errores de normalización
  }
}

// ==================== IMPORTACIÓN DESDE KOBO ====================
// Importación INCREMENTAL: solo agrega registros nuevos (por _uuid).
// Si DatosKobo no existe, hace importación completa inicial.
function importarCSVdesdeKobo() {
  try {
    var options = { 'muteHttpExceptions': true };
    var response = UrlFetchApp.fetch(URL_KOBO, options);
    var code = response.getResponseCode();

    if (code === 503) {
      SpreadsheetApp.getUi().alert(
        '⏳ KOBO ESTÁ OCUPADO (Código 503)\n\n' +
        'Espera 2 minutos y vuelve a intentarlo.'
      );
      return;
    }
    if (code !== 200) {
      SpreadsheetApp.getUi().alert('Error de Kobo (Código ' + code + '): ' + response.getContentText().substring(0, 200));
      return;
    }

    var csv = response.getContentText();
    var datosNuevos = Utilities.parseCsv(csv, ";");
    if (datosNuevos.length < 2) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Kobo no devolvió registros.', 'Aviso', 3);
      return;
    }

    var encNuevos = datosNuevos[0];
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = spreadsheet.getSheetByName("DatosKobo");

    // ── CASO 1: Primera importación (hoja no existe) ─────────────────
    if (!hoja) {
      hoja = spreadsheet.insertSheet("DatosKobo");
      hoja.getRange(1, 1, datosNuevos.length, encNuevos.length).setValues(datosNuevos);
      hoja.getRange(1, 1, 1, encNuevos.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
      hoja.setFrozenRows(1);
      limpiarColumnasKobo(hoja, encNuevos);
      normalizarAccionSilencioso(hoja);
      SpreadsheetApp.getActiveSpreadsheet().toast(
        '✅ Importación inicial: ' + (datosNuevos.length - 1) + ' registros normalizados', 'DatosKobo creado', 4);
      return;
    }

    // ── CASO 2: Importación incremental ──────────────────────────────
    // Detectar columna _uuid en el CSV nuevo
    var uuidColNuevo = -1;
    for (var i = 0; i < encNuevos.length; i++) {
      if (String(encNuevos[i]).trim().toLowerCase() === '_uuid') { uuidColNuevo = i; break; }
    }

    // Leer UUIDs ya existentes en DatosKobo
    var datosExist = hoja.getDataRange().getValues();
    var encExist = datosExist[0];
    var uuidColExist = -1;
    for (var i = 0; i < encExist.length; i++) {
      if (String(encExist[i]).trim().toLowerCase() === '_uuid') { uuidColExist = i; break; }
    }

    var uuidsExistentes = {};
    if (uuidColExist !== -1) {
      for (var f = 1; f < datosExist.length; f++) {
        var uid = String(datosExist[f][uuidColExist] || '').trim();
        if (uid) uuidsExistentes[uid] = true;
      }
    }

    // Filtrar solo filas nuevas
    var filasNuevas = [];
    for (var f = 1; f < datosNuevos.length; f++) {
      var uid = uuidColNuevo !== -1 ? String(datosNuevos[f][uuidColNuevo] || '').trim() : '';
      if (!uid || !uuidsExistentes[uid]) {
        filasNuevas.push(datosNuevos[f]);
      }
    }

    if (filasNuevas.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast('✅ Ya está al día. No hay registros nuevos.', 'Sin cambios', 4);
      return;
    }

    // Agregar filas nuevas al final
    var ultimaFila = hoja.getLastRow();
    hoja.getRange(ultimaFila + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);

    // Auto-normalizar: 🟢 Entrada / 🔴 Salida en la columna de acción
    normalizarAccionSilencioso(hoja);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '✅ ' + filasNuevas.length + ' registros nuevos añadidos y normalizados (total: ' + (datosExist.length - 1 + filasNuevas.length) + ')',
      'Actualización incremental', 4);

  } catch (e) {
    SpreadsheetApp.getUi().alert('Error al importar datos: ' + e.message);
  }
}

// Reimportación COMPLETA (borra DatosKobo y lo reconstruye desde cero)
// Usar solo cuando sea necesario reiniciar todo.
function reimportarTodoDesdKobo() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '⚠️ REIMPORTACIÓN COMPLETA',
    'Esto BORRARÁ DatosKobo y lo reimportará desde cero.\n\n' +
    'Usa esto solo si algo está muy desordenado.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    var options = { 'muteHttpExceptions': true };
    var response = UrlFetchApp.fetch(URL_KOBO, options);
    var code = response.getResponseCode();
    if (code !== 200) { ui.alert('Error Kobo (código ' + code + ')'); return; }

    var datos = Utilities.parseCsv(response.getContentText(), ";");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hoja = spreadsheet.getSheetByName("DatosKobo");
    if (hoja) spreadsheet.deleteSheet(hoja);
    hoja = spreadsheet.insertSheet("DatosKobo");
    hoja.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
    hoja.getRange(1, 1, 1, datos[0].length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
    limpiarColumnasKobo(hoja, datos[0]);
    ui.alert('✅ Reimportación completa: ' + (datos.length - 1) + ' registros.');
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

// Oculta columnas innecesarias de DatosKobo para mejor visualización
function limpiarColumnasKobo(hoja, encabezados) {
  // Columnas que queremos MOSTRAR (las importantes)
  var columnasImportantes = ['start', 'end', 'ingreso', 'egreso', 'entrada', 'salida',
    'participante', 'nombre', 'seleccione', 'c_id', '_uuid', 'uuid', 'accion', 'acción',
    'terapia', 'permiso', 'comput'];

  for (var i = 0; i < encabezados.length; i++) {
    var h = String(encabezados[i]).trim().toLowerCase();
    if (!h) {
      // Columna vacía - ocultar
      try { hoja.hideColumns(i + 1); } catch(e) {}
      continue;
    }

    var esImportante = false;
    for (var j = 0; j < columnasImportantes.length; j++) {
      if (h.indexOf(columnasImportantes[j]) !== -1 || h === columnasImportantes[j]) {
        esImportante = true;
        break;
      }
    }

    if (!esImportante) {
      try { hoja.hideColumns(i + 1); } catch(e) {}
    } else {
      try { hoja.showColumns(i + 1); } catch(e) {}
    }
  }
}

// Configurar triggers automáticos (cada hora + al abrir)
function configurarActualizacionAutomatica() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();
    if (handler === 'importarCSVdesdeKobo' || handler === 'importarAlAbrir') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // Trigger cada hora
  ScriptApp.newTrigger('importarCSVdesdeKobo')
    .timeBased()
    .everyHours(1)
    .create();

  // Trigger instalable al abrir (tiene permisos completos, a diferencia del onOpen simple)
  ScriptApp.newTrigger('importarAlAbrir')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();

  SpreadsheetApp.getUi().alert('✅ Actualización automática configurada:\n- Cada hora\n- Al abrir la hoja de cálculo');
}

// Función llamada por el trigger instalable onOpen (tiene permisos completos)
function importarAlAbrir() {
  importarCSVdesdeKobo();
}

// ==================== INSTALAR TODO ====================
function instalarTodo() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.alert(
    '🚀 INSTALAR TODO EL SISTEMA',
    'Esto ejecutará los siguientes pasos automáticamente:\n\n' +
    '1. Importar datos desde KoboToolbox\n' +
    '2. Diagnosticar columnas del CSV\n' +
    '3. Normalizar nombres de participantes\n' +
    '4. Crear hoja de Días de Estudio\n' +
    '5. Crear hoja de Lista de Terapias\n' +
    '6. Configurar actualización automática\n\n' +
    '¿Deseas continuar?',
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) return;

  var log = [];
  var errores = [];

  // --- PASO 1: Importar datos desde Kobo ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 1/6: Importando datos de Kobo...', '🚀 Instalando', -1);
    var options = { 'muteHttpExceptions': true };
    var response = UrlFetchApp.fetch(URL_KOBO, options);
    var code = response.getResponseCode();

    if (code === 200) {
      var csv = response.getContentText();
      var datos = Utilities.parseCsv(csv, ";");
      var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var hoja = spreadsheet.getSheetByName("DatosKobo");
      if (!hoja) hoja = spreadsheet.insertSheet("DatosKobo");
      hoja.clearContents();
      if (datos.length > 0) {
        hoja.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
        hoja.getRange(1, 1, 1, datos[0].length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
        hoja.setFrozenRows(1);
        limpiarColumnasKobo(hoja, datos[0]);
      }
      log.push('✅ Datos importados: ' + (datos.length - 1) + ' registros');
    } else if (code === 503) {
      log.push('⏳ Kobo ocupado (503) - los datos no se importaron. Reintenta más tarde.');
    } else {
      errores.push('❌ Error importando Kobo (código ' + code + ')');
    }
  } catch (e) {
    errores.push('❌ Error importando: ' + e.message);
  }

  // --- PASO 2: Diagnosticar columnas ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 2/6: Diagnosticando columnas...', '🚀 Instalando', -1);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaKobo = spreadsheet.getSheetByName("DatosKobo");
    if (hojaKobo) {
      var datosKobo = hojaKobo.getDataRange().getValues();
      var cols = detectarColumnas(datosKobo[0], datosKobo.slice(1));

      var colsInfo = [];
      if (cols.start !== undefined) colsInfo.push('start');
      if (cols.participante !== undefined) colsInfo.push('participante');
      if (cols.accionUnificada !== undefined) colsInfo.push('acción ("' + datosKobo[0][cols.accionUnificada] + '")');
      if (cols.ingreso !== undefined) colsInfo.push('ingreso (viejo)');
      if (cols.egreso !== undefined) colsInfo.push('egreso (viejo)');
      if (cols.uuid !== undefined) colsInfo.push('uuid');

      // Contar tipos
      var conteo = { entrada: 0, salida: 0, sinTipo: 0 };
      for (var f = 1; f < datosKobo.length; f++) {
        var t = obtenerTipoRegistro(datosKobo[f], cols);
        if (t.esIngreso) conteo.entrada++;
        else if (t.esEgreso) conteo.salida++;
        else conteo.sinTipo++;
      }

      log.push('✅ Columnas detectadas: ' + colsInfo.join(', '));
      log.push('   Entradas: ' + conteo.entrada + ' | Salidas: ' + conteo.salida + (conteo.sinTipo > 0 ? ' | Sin tipo: ' + conteo.sinTipo : ''));

      if (conteo.entrada === 0 && conteo.salida === 0) {
        errores.push('⚠️ No se detectaron entradas ni salidas. Revisa el formato del CSV con "Diagnosticar Datos Kobo".');
      }
    } else {
      errores.push('⚠️ No hay hoja DatosKobo - la importación pudo haber fallado');
    }
  } catch (e) {
    errores.push('❌ Error diagnosticando: ' + e.message);
  }

  // --- PASO 3: Normalizar nombres ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 3/6: Normalizando nombres...', '🚀 Instalando', -1);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var hojaKobo = spreadsheet.getSheetByName("DatosKobo");
    if (hojaKobo) {
      var datosKobo = hojaKobo.getDataRange().getValues();
      var colsP = buscarColumnasParticipante(datosKobo[0]);

      if (colsP.col1 !== -1) {
        // Recopilar nombres únicos (usando ambas columnas)
        var nombresUnicos = {};
        for (var f = 1; f < datosKobo.length; f++) {
          var nombre = obtenerNombreDeFila(datosKobo[f], colsP);
          if (nombre) nombresUnicos[nombre] = true;
        }
        var todosNombres = Object.keys(nombresUnicos);

        // Agrupar por código
        var porCodigo = {};
        var sinCodigo = [];
        for (var i = 0; i < todosNombres.length; i++) {
          var n = todosNombres[i];
          var codigo = extraerCodigo(n);
          if (codigo) {
            if (!porCodigo[codigo]) porCodigo[codigo] = [];
            porCodigo[codigo].push(n);
          } else {
            sinCodigo.push(n);
          }
        }

        // Emparejar nombres sin código
        for (var i = 0; i < sinCodigo.length; i++) {
          var nombreSinCod = sinCodigo[i];
          var limpio = textoParaComparar(limpiarNombre(nombreSinCod));
          var encontrado = false;
          var codigos = Object.keys(porCodigo);
          for (var c = 0; c < codigos.length; c++) {
            var grupo = porCodigo[codigos[c]];
            for (var g = 0; g < grupo.length; g++) {
              var limpioGrupo = textoParaComparar(limpiarNombre(grupo[g]));
              if (nombresCoinciden(limpio, limpioGrupo)) {
                grupo.push(nombreSinCod);
                encontrado = true;
                break;
              }
            }
            if (encontrado) break;
          }
          if (!encontrado) porCodigo['_SIN_' + i] = [nombreSinCod];
        }

        // Generar mapeo
        var mapeo = {};
        var codigosKeys = Object.keys(porCodigo);
        for (var c = 0; c < codigosKeys.length; c++) {
          var grupo = porCodigo[codigosKeys[c]];
          var codigoReal = codigosKeys[c].indexOf('_SIN_') === 0 ? '' : codigosKeys[c];
          var mejorNombre = '';
          var mejorLargo = 0;
          for (var g = 0; g < grupo.length; g++) {
            var limpio = limpiarNombre(grupo[g]);
            if (limpio.length > mejorLargo) { mejorLargo = limpio.length; mejorNombre = limpio; }
          }
          var canonico = codigoReal ? mejorNombre + ' (' + codigoReal + ')' : mejorNombre;
          for (var g = 0; g < grupo.length; g++) { mapeo[grupo[g]] = canonico; }
        }

        // Preservar correcciones manuales
        var hojaNombres = spreadsheet.getSheetByName("NombresCanonicos");
        if (hojaNombres) {
          var datosExistentes = hojaNombres.getDataRange().getValues();
          for (var f = 1; f < datosExistentes.length; f++) {
            var orig = String(datosExistentes[f][0] || '').trim();
            var canon = String(datosExistentes[f][1] || '').trim();
            if (orig && canon && mapeo[orig]) mapeo[orig] = canon;
          }
        }

        // Escribir hoja NombresCanonicos
        if (!hojaNombres) hojaNombres = spreadsheet.insertSheet("NombresCanonicos");
        else hojaNombres.clearContents();

        var encNombres = ['Nombre Original (Kobo)', 'Nombre Canónico', 'Código'];
        hojaNombres.getRange(1, 1, 1, 3).setValues([encNombres]);
        hojaNombres.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#ff6f00').setFontColor('#ffffff').setHorizontalAlignment('center');
        hojaNombres.setFrozenRows(1);

        var filas = [];
        var claves = Object.keys(mapeo).sort();
        var duplicados = 0;
        for (var i = 0; i < claves.length; i++) {
          var cod = extraerCodigo(claves[i]) || '';
          filas.push([claves[i], mapeo[claves[i]], cod]);
          if (claves[i] !== mapeo[claves[i]]) duplicados++;
        }
        if (filas.length > 0) hojaNombres.getRange(2, 1, filas.length, 3).setValues(filas);
        hojaNombres.setColumnWidth(1, 350);
        hojaNombres.setColumnWidth(2, 350);
        hojaNombres.setColumnWidth(3, 130);

        for (var i = 0; i < filas.length; i++) {
          if (filas[i][0] !== filas[i][1]) hojaNombres.getRange(i + 2, 1, 1, 3).setBackground('#fff3e0');
        }

        // Normalizar en otras hojas
        normalizarNombresEnHojas(mapeo, spreadsheet);

        log.push('✅ Nombres normalizados: ' + claves.length + ' únicos, ' + duplicados + ' agrupados');
      }
    }
  } catch (e) {
    errores.push('❌ Error normalizando: ' + e.message);
  }

  // --- PASO 4: Crear hoja DiasEstudio ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 4/6: Configurando Días de Estudio...', '🚀 Instalando', -1);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet.getSheetByName("DiasEstudio")) {
      crearHojaDiasEstudio();
      log.push('✅ Hoja DiasEstudio creada');
    } else {
      log.push('ℹ️ Hoja DiasEstudio ya existe (no se tocó)');
    }
  } catch (e) {
    errores.push('❌ Error en DiasEstudio: ' + e.message);
  }

  // --- PASO 5: Crear hoja ListaTerapias ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 5/6: Configurando Lista de Terapias...', '🚀 Instalando', -1);
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet.getSheetByName("ListaTerapias")) {
      crearHojaListaTerapias();
      log.push('✅ Hoja ListaTerapias creada');
    } else {
      log.push('ℹ️ Hoja ListaTerapias ya existe (no se tocó)');
    }
  } catch (e) {
    errores.push('❌ Error en ListaTerapias: ' + e.message);
  }

  // --- PASO 6: Configurar triggers automáticos ---
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Paso 6/6: Configurando triggers automáticos...', '🚀 Instalando', -1);
    // Limpiar triggers existentes
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      var handler = triggers[i].getHandlerFunction();
      if (handler === 'importarCSVdesdeKobo' || handler === 'importarAlAbrir') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    ScriptApp.newTrigger('importarCSVdesdeKobo').timeBased().everyHours(1).create();
    ScriptApp.newTrigger('importarAlAbrir').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
    log.push('✅ Actualización automática: cada hora + al abrir');
  } catch (e) {
    errores.push('❌ Error en triggers: ' + e.message);
  }

  // --- RESUMEN FINAL ---
  SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1); // Cerrar toast
  var resumen = '🚀 INSTALACIÓN COMPLETA\n\n';
  resumen += log.join('\n') + '\n';
  if (errores.length > 0) {
    resumen += '\n--- PROBLEMAS ---\n' + errores.join('\n') + '\n';
  }
  resumen += '\n--- HOJAS DEL SISTEMA ---\n';
  resumen += '• DatosKobo - datos crudos de Kobo\n';
  resumen += '• NombresCanonicos - mapeo de nombres (editable)\n';
  resumen += '• DiasEstudio - marca días de estudio por persona\n';
  resumen += '• ListaTerapias - marca quién recibe terapia\n';
  resumen += '\nYa puedes generar reportes desde el menú.';

  ui.alert(resumen);
}

// ==================== DÍAS DE ESTUDIO ====================
// Crea o abre la hoja "DiasEstudio" donde se configuran los días que cada participante estudia
function crearHojaDiasEstudio() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("DiasEstudio");

  if (!hoja) {
    hoja = spreadsheet.insertSheet("DiasEstudio");

    // Encabezados
    var encabezados = ['Participante', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
    hoja.getRange(1, 1, 1, 8).setValues([encabezados]);
    hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#7b1fa2').setFontColor('#ffffff').setHorizontalAlignment('center');
    hoja.setFrozenRows(1);

    // Llenar con participantes existentes si hay datos de Kobo
    var hojaKobo = spreadsheet.getSheetByName("DatosKobo");
    if (hojaKobo) {
      var datosKobo = hojaKobo.getDataRange().getValues();
      var colsP = buscarColumnasParticipante(datosKobo[0]);
      if (colsP.col1 !== -1) {
        var mapeoNombres = cargarMapeoNombres();
        var participantes = {};
        for (var f = 1; f < datosKobo.length; f++) {
          var nombre = obtenerNombreDeFila(datosKobo[f], colsP);
          if (nombre) participantes[normalizarNombre(nombre, mapeoNombres)] = true;
        }
        var lista = Object.keys(participantes).sort();
        for (var p = 0; p < lista.length; p++) {
          hoja.getRange(p + 2, 1).setValue(lista[p]);
        }
      }
    }

    // Validación: solo permitir "X" o vacío en las columnas de días
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['X', ''], true)
      .setAllowInvalid(false)
      .setHelpText('Escribe X si ese día es de estudio, déjalo vacío si no')
      .build();
    hoja.getRange(2, 2, 50, 7).setDataValidation(regla);

    // Ancho de columnas
    hoja.setColumnWidth(1, 200);
    for (var c = 2; c <= 8; c++) {
      hoja.setColumnWidth(c, 100);
    }

    hoja.getRange(2, 2, 50, 7).setHorizontalAlignment('center');

    SpreadsheetApp.getUi().alert(
      '📚 HOJA DE DÍAS DE ESTUDIO CREADA\n\n' +
      'Instrucciones:\n' +
      '1. En la columna "Participante" escribe el nombre exacto como aparece en Kobo\n' +
      '2. Marca con "X" los días que esa persona tiene clase/estudio\n' +
      '3. Los días marcados aparecerán como "Día de Estudio" en el reporte con 0% de pago\n\n' +
      'Ejemplo: Si "Juan" estudia los Martes y Jueves, pon X en esas columnas'
    );
  }

  hoja.activate();
}

// Lee la hoja DiasEstudio y devuelve un mapa: { "participante": [0,1,0,1,0,0,0] } (Lun-Dom)
function obtenerDiasEstudio() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("DiasEstudio");
  var mapa = {};

  if (!hoja) return mapa;

  var datos = hoja.getDataRange().getValues();
  // Columnas: A=Participante, B-H=Lun-Dom, I=Fecha Inicio, J=Fecha Fin
  for (var f = 1; f < datos.length; f++) {
    var participante = String(datos[f][0] || '').trim();
    if (!participante) continue;

    // Columnas 1-7 corresponden a Lunes(1), Martes(2), Miércoles(3), Jueves(4), Viernes(5), Sábado(6), Domingo(0)
    // Mapeamos a índice JS de día de semana: 0=Domingo, 1=Lunes, ..., 6=Sábado
    var dias = {};
    dias[1] = String(datos[f][1] || '').trim().toUpperCase() === 'X'; // Lunes
    dias[2] = String(datos[f][2] || '').trim().toUpperCase() === 'X'; // Martes
    dias[3] = String(datos[f][3] || '').trim().toUpperCase() === 'X'; // Miércoles
    dias[4] = String(datos[f][4] || '').trim().toUpperCase() === 'X'; // Jueves
    dias[5] = String(datos[f][5] || '').trim().toUpperCase() === 'X'; // Viernes
    dias[6] = String(datos[f][6] || '').trim().toUpperCase() === 'X'; // Sábado
    dias[0] = String(datos[f][7] || '').trim().toUpperCase() === 'X'; // Domingo

    // Fechas de vigencia (columnas I y J) — opcionales
    var fechaInicio = datos[f][8] ? new Date(datos[f][8]) : null;
    var fechaFin = datos[f][9] ? new Date(datos[f][9]) : null;

    // Si hay fecha fin, ajustar al final del día
    if (fechaFin) {
      fechaFin.setHours(23, 59, 59, 999);
    }
    // Si hay fecha inicio, ajustar al inicio del día
    if (fechaInicio) {
      fechaInicio.setHours(0, 0, 0, 0);
    }

    mapa[participante] = {
      dias: dias,
      fechaInicio: fechaInicio,
      fechaFin: fechaFin
    };
  }

  return mapa;
}

// Verifica si una fecha es día de estudio para un participante
function esDiaDeEstudio(participante, fecha, diasEstudioMapa) {
  if (!diasEstudioMapa[participante]) return false;

  var config = diasEstudioMapa[participante];

  // Verificar si la fecha está dentro del rango de vigencia
  if (config.fechaInicio && fecha < config.fechaInicio) return false;
  if (config.fechaFin && fecha > config.fechaFin) return false;

  var diaSemana = fecha.getDay(); // 0=Domingo, 1=Lunes, ...
  return config.dias[diaSemana] === true;
}

// ==================== LISTA DE TERAPIAS ====================
// Crea o abre la hoja "ListaTerapias" donde se registran las personas que asisten a terapia
function crearHojaListaTerapias() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("ListaTerapias");

  if (!hoja) {
    hoja = spreadsheet.insertSheet("ListaTerapias");

    // Encabezados
    var encabezados = ['Participante', 'Recibe Terapia (X)', 'Notas'];
    hoja.getRange(1, 1, 1, 3).setValues([encabezados]);
    hoja.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#00897b').setFontColor('#ffffff').setHorizontalAlignment('center');
    hoja.setFrozenRows(1);

    // Llenar con participantes existentes
    var hojaKobo = spreadsheet.getSheetByName("DatosKobo");
    if (hojaKobo) {
      var datosKobo = hojaKobo.getDataRange().getValues();
      var colsP = buscarColumnasParticipante(datosKobo[0]);
      if (colsP.col1 !== -1) {
        var mapeoNombres = cargarMapeoNombres();
        var participantes = {};
        for (var f = 1; f < datosKobo.length; f++) {
          var nombre = obtenerNombreDeFila(datosKobo[f], colsP);
          if (nombre) participantes[normalizarNombre(nombre, mapeoNombres)] = true;
        }
        var lista = Object.keys(participantes).sort();
        for (var p = 0; p < lista.length; p++) {
          hoja.getRange(p + 2, 1).setValue(lista[p]);
        }
      }
    }

    // Validación
    var regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['X', ''], true)
      .build();
    hoja.getRange(2, 2, 100, 1).setDataValidation(regla);
    hoja.getRange(2, 2, 100, 1).setHorizontalAlignment('center');

    hoja.setColumnWidth(1, 250);
    hoja.setColumnWidth(2, 150);
    hoja.setColumnWidth(3, 300);

    SpreadsheetApp.getUi().alert('🧘 HOJA DE LISTA DE TERAPIAS CREADA\n\nMarca con una "X" a las personas que asisten a terapias.');
  }
  hoja.activate();
}

// Obtiene el set de personas que reciben terapia
function obtenerListaTerapias() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("ListaTerapias");
  var lista = {};

  if (!hoja) return lista;

  var datos = hoja.getDataRange().getValues();
  for (var f = 1; f < datos.length; f++) {
    var participante = String(datos[f][0] || '').trim();
    var recibe = String(datos[f][1] || '').trim().toUpperCase() === 'X';
    if (participante && recibe) {
      lista[participante] = true;
    }
  }
  return lista;
}

// ==================== NORMALIZACIÓN DE NOMBRES ====================
// Extrae el código de participante de un nombre (ej. "VILÓ040971" de "VILÓ040971 Vilma López" o "Vilma Lopez (VILÓ040971)")
function extraerCodigo(nombre) {
  // Patrón: 4 letras mayúsculas (pueden incluir acentos) + 6 dígitos
  var match = nombre.match(/([A-ZÁÉÍÓÚÑÜ]{4}\d{6})/i);
  return match ? match[1].toUpperCase() : null;
}

// Quita el código del nombre para obtener solo el nombre limpio
function limpiarNombre(nombre) {
  // Quitar código al inicio: "VILÓ040971 Vilma López" -> "Vilma López"
  var sinCodigo = nombre.replace(/^[A-ZÁÉÍÓÚÑÜ]{4}\d{6}\s*/i, '');
  // Quitar código entre paréntesis: "Vilma Lopez (VILÓ040971)" -> "Vilma Lopez"
  sinCodigo = sinCodigo.replace(/\s*\([A-ZÁÉÍÓÚÑÜ]{4}\d{6}\)\s*/i, '');
  // Quitar bullets y espacios extra
  sinCodigo = sinCodigo.replace(/^[•\s]+/, '').trim();
  return sinCodigo;
}

// Normaliza texto para comparación (sin acentos, minúsculas, sin espacios dobles)
function textoParaComparar(texto) {
  return texto.toLowerCase()
    .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n')
    .replace(/\s+/g, ' ').trim();
}

// Compara dos nombres normalizados para ver si son la misma persona
// Ej: "juana vicente" y "juana del rosario vicente choy" → true (ambas palabras del corto están en el largo)
function nombresCoinciden(nombre1, nombre2) {
  if (nombre1 === nombre2) return true;
  if (nombre1.indexOf(nombre2) !== -1 || nombre2.indexOf(nombre1) !== -1) return true;

  // Verificar si TODAS las palabras significativas del nombre más corto están en el más largo
  var corto = nombre1.length <= nombre2.length ? nombre1 : nombre2;
  var largo = nombre1.length <= nombre2.length ? nombre2 : nombre1;
  var palabras = corto.split(' ');
  var coincidencias = 0;
  var totalSignificativas = 0;

  for (var i = 0; i < palabras.length; i++) {
    // Ignorar palabras muy cortas (de, del, la, el, etc.)
    if (palabras[i].length < 3) continue;
    totalSignificativas++;
    if (largo.indexOf(palabras[i]) !== -1) coincidencias++;
  }

  // Si todas las palabras significativas del corto están en el largo, es la misma persona
  if (totalSignificativas > 0 && coincidencias === totalSignificativas) return true;

  // Si al menos 2 palabras significativas coinciden y es más del 60%, también
  if (coincidencias >= 2 && totalSignificativas > 0 && (coincidencias / totalSignificativas) >= 0.6) return true;

  return false;
}

// Crea/actualiza la hoja NombresCanonicos con el mapeo automático
function crearHojaNombresCanonicos() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = spreadsheet.getSheetByName("DatosKobo");

  if (!hojaKobo) {
    SpreadsheetApp.getUi().alert('Primero importa datos desde KoboToolbox.');
    return;
  }

  var datosKobo = hojaKobo.getDataRange().getValues();
  var colsP = buscarColumnasParticipante(datosKobo[0]);
  if (colsP.col1 === -1) {
    SpreadsheetApp.getUi().alert('No se encontró la columna de participantes en DatosKobo.');
    return;
  }

  // Recopilar todos los nombres únicos (usando ambas columnas)
  var nombresUnicos = {};
  for (var f = 1; f < datosKobo.length; f++) {
    var nombre = obtenerNombreDeFila(datosKobo[f], colsP);
    if (nombre) nombresUnicos[nombre] = true;
  }
  var todosNombres = Object.keys(nombresUnicos);

  // Agrupar por código
  var porCodigo = {};   // código -> [nombres que tienen ese código]
  var sinCodigo = [];   // nombres sin código

  for (var i = 0; i < todosNombres.length; i++) {
    var n = todosNombres[i];
    var codigo = extraerCodigo(n);
    if (codigo) {
      if (!porCodigo[codigo]) porCodigo[codigo] = [];
      porCodigo[codigo].push(n);
    } else {
      sinCodigo.push(n);
    }
  }

  // Para nombres sin código, intentar emparejar con grupos existentes por similitud de nombre
  for (var i = 0; i < sinCodigo.length; i++) {
    var nombreSinCod = sinCodigo[i];
    var limpio = textoParaComparar(limpiarNombre(nombreSinCod));
    var encontrado = false;

    var codigos = Object.keys(porCodigo);
    for (var c = 0; c < codigos.length; c++) {
      var grupo = porCodigo[codigos[c]];
      for (var g = 0; g < grupo.length; g++) {
        var limpioGrupo = textoParaComparar(limpiarNombre(grupo[g]));
        // Coincidencia: uno contiene al otro, o son muy similares
        if (nombresCoinciden(limpio, limpioGrupo)) {
          grupo.push(nombreSinCod);
          encontrado = true;
          break;
        }
      }
      if (encontrado) break;
    }

    if (!encontrado) {
      // Queda como grupo propio sin código
      porCodigo['_SIN_' + i] = [nombreSinCod];
    }
  }

  // Para cada grupo, elegir el nombre canónico: el nombre completo más largo (sin código)
  var mapeo = {}; // nombre original -> nombre canónico
  var codigos = Object.keys(porCodigo);
  for (var c = 0; c < codigos.length; c++) {
    var grupo = porCodigo[codigos[c]];
    var codigoReal = codigos[c].indexOf('_SIN_') === 0 ? '' : codigos[c];

    // Elegir el nombre más largo (limpio) como canónico
    var mejorNombre = '';
    var mejorLargo = 0;
    for (var g = 0; g < grupo.length; g++) {
      var limpio = limpiarNombre(grupo[g]);
      if (limpio.length > mejorLargo) {
        mejorLargo = limpio.length;
        mejorNombre = limpio;
      }
    }

    // Formato canónico: "Nombre Completo (CÓDIGO)" o solo "Nombre Completo"
    var canonico = codigoReal ? mejorNombre + ' (' + codigoReal + ')' : mejorNombre;

    for (var g = 0; g < grupo.length; g++) {
      mapeo[grupo[g]] = canonico;
    }
  }

  // Leer mapeo existente para preservar correcciones manuales del usuario
  var hoja = spreadsheet.getSheetByName("NombresCanonicos");
  var mapeoManual = {};
  if (hoja) {
    var datosExistentes = hoja.getDataRange().getValues();
    for (var f = 1; f < datosExistentes.length; f++) {
      var orig = String(datosExistentes[f][0] || '').trim();
      var canon = String(datosExistentes[f][1] || '').trim();
      if (orig && canon) mapeoManual[orig] = canon;
    }
  }

  // Combinar: mapeo manual tiene prioridad
  var keys = Object.keys(mapeo);
  for (var k = 0; k < keys.length; k++) {
    if (mapeoManual[keys[k]]) {
      mapeo[keys[k]] = mapeoManual[keys[k]];
    }
  }

  // Crear/actualizar hoja
  if (!hoja) {
    hoja = spreadsheet.insertSheet("NombresCanonicos");
  } else {
    hoja.clearContents();
  }

  var encabezados = ['Nombre Original (Kobo)', 'Nombre Canónico', 'Código'];
  hoja.getRange(1, 1, 1, 3).setValues([encabezados]);
  hoja.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#ff6f00').setFontColor('#ffffff').setHorizontalAlignment('center');
  hoja.setFrozenRows(1);

  var filas = [];
  var claves = Object.keys(mapeo).sort();
  for (var i = 0; i < claves.length; i++) {
    var codigo = extraerCodigo(claves[i]) || '';
    filas.push([claves[i], mapeo[claves[i]], codigo]);
  }

  if (filas.length > 0) {
    hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  }

  hoja.setColumnWidth(1, 350);
  hoja.setColumnWidth(2, 350);
  hoja.setColumnWidth(3, 130);

  // Resaltar filas donde original != canónico (duplicados detectados)
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] !== filas[i][1]) {
      hoja.getRange(i + 2, 1, 1, 3).setBackground('#fff3e0');
    }
  }

  hoja.activate();

  var duplicados = 0;
  for (var i = 0; i < filas.length; i++) {
    if (filas[i][0] !== filas[i][1]) duplicados++;
  }

  // Normalizar nombres en las hojas DiasEstudio y ListaTerapias también
  var hojasActualizadas = normalizarNombresEnHojas(mapeo, spreadsheet);

  SpreadsheetApp.getUi().alert(
    '🔀 NOMBRES NORMALIZADOS\n\n' +
    'Se encontraron ' + claves.length + ' nombres únicos.\n' +
    duplicados + ' variantes fueron agrupadas automáticamente.\n\n' +
    'Revisa la hoja "NombresCanonicos":\n' +
    '- Columna A: nombre tal como viene de Kobo\n' +
    '- Columna B: nombre canónico (puedes editarlo)\n' +
    '- Las filas en naranja son las que se unificaron\n\n' +
    (hojasActualizadas ? 'También se actualizaron: ' + hojasActualizadas + '\n\n' : '') +
    'Los reportes usarán automáticamente los nombres canónicos.'
  );
}

// Normaliza los nombres en DiasEstudio y ListaTerapias usando el mapeo
function normalizarNombresEnHojas(mapeo, spreadsheet) {
  var hojasActualizadas = [];

  // Actualizar DiasEstudio
  var hojaDias = spreadsheet.getSheetByName("DiasEstudio");
  if (hojaDias) {
    var datos = hojaDias.getDataRange().getValues();
    var cambios = 0;
    for (var f = 1; f < datos.length; f++) {
      var nombre = String(datos[f][0] || '').trim();
      if (nombre && mapeo[nombre] && mapeo[nombre] !== nombre) {
        hojaDias.getRange(f + 1, 1).setValue(mapeo[nombre]);
        cambios++;
      }
    }
    if (cambios > 0) hojasActualizadas.push('DiasEstudio (' + cambios + ' nombres)');
  }

  // Actualizar ListaTerapias
  var hojaTerapias = spreadsheet.getSheetByName("ListaTerapias");
  if (hojaTerapias) {
    var datos = hojaTerapias.getDataRange().getValues();
    var cambios = 0;
    for (var f = 1; f < datos.length; f++) {
      var nombre = String(datos[f][0] || '').trim();
      if (nombre && mapeo[nombre] && mapeo[nombre] !== nombre) {
        hojaTerapias.getRange(f + 1, 1).setValue(mapeo[nombre]);
        cambios++;
      }
    }
    if (cambios > 0) hojasActualizadas.push('ListaTerapias (' + cambios + ' nombres)');
  }

  return hojasActualizadas.length > 0 ? hojasActualizadas.join(', ') : '';
}

// Carga el mapeo de normalización para usar en reportes
function cargarMapeoNombres() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = spreadsheet.getSheetByName("NombresCanonicos");
  var mapa = {};

  if (!hoja) return mapa;

  var datos = hoja.getDataRange().getValues();
  for (var f = 1; f < datos.length; f++) {
    var original = String(datos[f][0] || '').trim();
    var canonico = String(datos[f][1] || '').trim();
    if (original && canonico) {
      mapa[original] = canonico;
    }
  }
  return mapa;
}

// Aplica normalización a un nombre
function normalizarNombre(nombre, mapeoNombres) {
  if (!mapeoNombres || Object.keys(mapeoNombres).length === 0) return nombre;
  return mapeoNombres[nombre] || nombre;
}

// ==================== FUNCIONES DE INTERFAZ ====================
function solicitarParametrosReporte() {
  var ui = SpreadsheetApp.getUi();
  
  // 1. Preguntar por Participante
  var participantes = obtenerListaParticipantes();
  var promptPart = ui.prompt('👤 REPORTE INDIVIDUAL', 
    'Escribe el nombre del participante (exacto) o deja vacío para TODOS.\n\nSugerencias:\n' + participantes.slice(0, 10).join('\n'), 
    ui.ButtonSet.OK_CANCEL);
  
  if (promptPart.getSelectedButton() !== ui.Button.OK) return null;
  var filtro = promptPart.getResponseText().trim();
  
  // 2. Preguntar por Nueva Hoja
  var resHoja = ui.alert('📄 ¿CREAR NUEVA HOJA?', 
    '¿Deseas crear una nueva pestaña para este reporte?\n\nSi eliges NO, se usará la hoja "Reporte_Fijo" (se sobreescribirá).', 
    ui.ButtonSet.YES_NO);
    
  var nuevaHoja = (resHoja === ui.Button.YES);
  
  return { filtro: filtro, nuevaHoja: nuevaHoja };
}

function obtenerListaParticipantes() {
  var hojaKobo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DatosKobo");
  if (!hojaKobo) return [];

  var datos = hojaKobo.getDataRange().getValues();
  var colsP = buscarColumnasParticipante(datos[0]);
  if (colsP.col1 === -1) return [];

  var mapeoNombres = cargarMapeoNombres();
  var names = {};
  for (var f = 1; f < datos.length; f++) {
    var n = obtenerNombreDeFila(datos[f], colsP);
    if (n) {
      var normalizado = normalizarNombre(n, mapeoNombres);
      names[normalizado] = true;
    }
  }
  return Object.keys(names).sort();
}

// ==================== FUNCIONES DE REPORTES ====================
function generarReportePorDia() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('📅 REPORTE POR DÍA', 'Ingresa la fecha (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    var partes = respuesta.getResponseText().split('/');
    if (partes.length === 3) {
      var params = solicitarParametrosReporte();
      if (!params) return;
      var fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      generarReporte('dia', fecha, null, params.filtro, params.nuevaHoja);
    } else {
      ui.alert('Error: Formato incorrecto. Usa dd/mm/yyyy');
    }
  }
}

function generarReportePorSemana() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('📆 REPORTE POR SEMANA', 'Ingresa la fecha de inicio (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    var partes = respuesta.getResponseText().split('/');
    if (partes.length === 3) {
      var params = solicitarParametrosReporte();
      if (!params) return;
      var fecha = new Date(partes[2], partes[1] - 1, partes[0]);
      generarReporte('semana', fecha, null, params.filtro, params.nuevaHoja);
    } else {
      ui.alert('Error: Formato incorrecto. Usa dd/mm/yyyy');
    }
  }
}

function generarReportePorMes() {
  var ui = SpreadsheetApp.getUi();
  var respuesta = ui.prompt('🗓️ REPORTE POR MES', 'Ingresa mes y año (mm/yyyy):', ui.ButtonSet.OK_CANCEL);

  if (respuesta.getSelectedButton() == ui.Button.OK) {
    var partes = respuesta.getResponseText().split('/');
    if (partes.length === 2) {
      var params = solicitarParametrosReporte();
      if (!params) return;
      var fecha = new Date(partes[1], partes[0] - 1, 1);
      generarReporte('mes', fecha, null, params.filtro, params.nuevaHoja);
    } else {
      ui.alert('Error: Formato incorrecto. Usa mm/yyyy');
    }
  }
}

function generarReportePorRango() {
  var ui = SpreadsheetApp.getUi();
  var resInicio = ui.prompt('📊 REPORTE POR RANGO', 'Fecha Inicio (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (resInicio.getSelectedButton() != ui.Button.OK) return;
  var resFin = ui.prompt('📊 REPORTE POR RANGO', 'Fecha Fin (dd/mm/yyyy):', ui.ButtonSet.OK_CANCEL);
  if (resFin.getSelectedButton() != ui.Button.OK) return;

  var p1 = resInicio.getResponseText().split('/');
  var p2 = resFin.getResponseText().split('/');
  if (p1.length === 3 && p2.length === 3) {
    var params = solicitarParametrosReporte();
    if (!params) return;
    var f1 = new Date(p1[2], p1[1] - 1, p1[0]);
    var f2 = new Date(p2[2], p2[1] - 1, p2[0]);
    generarReporte('rango', f1, f2, params.filtro, params.nuevaHoja);
  } else {
    ui.alert('Error: Formato incorrecto.');
  }
}

function generarReporteTodo() {
  var params = solicitarParametrosReporte();
  if (!params) return;
  generarReporte('todo', null, null, params.filtro, params.nuevaHoja);
}

// ==================== FUNCIÓN PRINCIPAL - CRONOLÓGICA ====================
// ==================== FUNCIÓN PRINCIPAL - CRONOLÓGICA ====================
function generarReporte(tipo, fechaInicio, fechaFin, empleadoFiltro, nuevaHoja) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaActual = spreadsheet.getSheetByName("DatosKobo");

  if (!hojaActual) {
    SpreadsheetApp.getUi().alert('ERROR\n\nNo existe la hoja "DatosKobo".\nPrimero importa los datos desde KoboToolbox.');
    return;
  }

  var datos = hojaActual.getDataRange().getValues();
  var encabezados = datos[0];

  // 1. Identificar columnas - soporta ambos formatos de Kobo
  var cols = detectarColumnas(encabezados, datos.slice(1));

  if (cols.start === undefined || cols.participante === undefined) {
    SpreadsheetApp.getUi().alert('ERROR DE COLUMNAS\n\nNo se encontró "start" o "Participante".\nUsa "Diagnosticar Datos Kobo" para ver las columnas detectadas.');
    return;
  }
  if (cols.accionUnificada === undefined && cols.ingreso === undefined && cols.egreso === undefined) {
    SpreadsheetApp.getUi().alert('ERROR DE COLUMNAS\n\nNo se encontró columna de Ingreso/Egreso.\nUsa "Diagnosticar Datos Kobo" para ver las columnas detectadas.');
    return;
  }

  // 2. Filtrar Duplicados y Extraer Registros
  var uuidVistos = {};
  var registrosVistos = {}; // Backup por si no hay UUID

  // Cargar configuraciones
  var diasEstudioMapa = obtenerDiasEstudio();
  var listaTerapias = obtenerListaTerapias();
  var mapeoNombres = cargarMapeoNombres();
  var tz = Session.getScriptTimeZone();

  // Preparar hoja de reporte
  var periodo = obtenerTextoPeriodo(tipo, fechaInicio, fechaFin);
  var nombreHoja = nuevaHoja ? ('Reporte_' + Utilities.formatDate(new Date(), tz, 'ddMM_HHmm')) : 'Reporte_Fijo';
  var hoja = spreadsheet.getSheetByName(nombreHoja);
  
  if (hoja) {
    if (!nuevaHoja) {
      hoja.clear().activate();
    } else {
      hoja = spreadsheet.insertSheet(nombreHoja + '_' + new Date().getTime());
    }
  } else {
    hoja = spreadsheet.insertSheet(nombreHoja);
  }

  var encabezadosReporte = ['Fecha', 'Día', '🟢 Entrada', '🔴 Salida', 'Tipo', 'Horas Trabajadas', 'Porcentaje', 'Horas a Pagar'];
  hoja.getRange(1, 1, 1, 8).setValues([encabezadosReporte]);
  hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1f54a8').setFontColor('#ffffff').setHorizontalAlignment('center');
  hoja.setFrozenRows(1);

  var filaActual = 2;
  var tituloReporte = 'Período: ' + periodo + (empleadoFiltro ? ' | Filtro: ' + empleadoFiltro : ' | Todos');
  hoja.getRange(filaActual, 1, 1, 8).merge().setValue(tituloReporte);
  hoja.getRange(filaActual, 1).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setBackground('#e8eaf6');
  filaActual++;

  // 1. Agrupar y Limpiar Datos
  var registrosPorEmpleado = {};
  for (var f = 1; f < datos.length; f++) {
    var empOriginal = obtenerParticipanteFila(datos[f], cols);
    if (!empOriginal) continue;
    var emp = normalizarNombre(empOriginal, mapeoNombres);

    // --- FILTRO DE ENTRADAS DUPLICADAS (Sync errors de Kobo) ---
    if (cols.uuid !== undefined) {
      var uid = datos[f][cols.uuid];
      if (uid) {
        if (uuidVistos[uid]) continue;
        uuidVistos[uid] = true;
      }
    }
    var row = datos[f];

    // Determinar tipo de registro según formato del CSV
    var tipoReg = obtenerTipoRegistro(row, cols);

    // Saltar filas sin tipo (errores de datos en Kobo donde no seleccionaron Ingreso/Egreso)
    if (!tipoReg.esIngreso && !tipoReg.esEgreso) continue;

    // Fallback dedup por si no hay UUID
    var claveUnica = emp + '|' + row[cols.start] + '|' + tipoReg.tipo;
    if (registrosVistos[claveUnica]) continue;
    registrosVistos[claveUnica] = true;

    // FILTRO DE EMPLEADO (comparar con nombre normalizado)
    if (empleadoFiltro && emp.toLowerCase() !== empleadoFiltro.toLowerCase()) continue;

    var fechaStr = row[cols.start];
    if (!fechaStr) continue;
    var fechaObj = new Date(fechaStr);

    // Para Salidas: usar la columna 'end' como tiempo real de salida.
    // Razón: el trabajador abre el formulario (start) pero lo envía al salir (end).
    // Solo se aplica si 'end' existe, es posterior a 'start', y la diferencia es razonable (< 24h).
    if (tipoReg.esEgreso && cols.end !== undefined && row[cols.end]) {
      var fechaEnd = new Date(row[cols.end]);
      if (!isNaN(fechaEnd.getTime())) {
        var diffHoras = (fechaEnd - fechaObj) / (1000 * 60 * 60);
        if (diffHoras > 0 && diffHoras < 24) {
          fechaObj = fechaEnd;
        }
      }
    }

    if (!validarEnRango(tipo, fechaObj, fechaInicio, fechaFin)) continue;

    if (!registrosPorEmpleado[emp]) registrosPorEmpleado[emp] = [];
    registrosPorEmpleado[emp].push({
      fila: f,
      fecha: fechaObj,
      esIngreso: tipoReg.esIngreso,
      esEgreso: tipoReg.esEgreso,
      esTerapia: tipoReg.esTerapia,
      esPermiso: tipoReg.esPermiso,
      esComputacion: tipoReg.esComputacion
    });
  }

  var listaEmpleados = Object.keys(registrosPorEmpleado).sort();
  if (listaEmpleados.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontraron registros que coincidan con la búsqueda.');
    return;
  }
  
  var totalesGen = { horas: 0, pagar: 0, terapia: 0, personas: listaEmpleados.length };

  // 2. Procesar cada Empleado
  listaEmpleados.forEach(function(empleadoId) {
    var registros = registrosPorEmpleado[empleadoId];
    registros.sort(function(a, b) { return a.fecha - b.fecha; });

    filaActual++;
    hoja.getRange(filaActual, 1, 1, 8).merge().setValue('👤 ' + empleadoId.toUpperCase());
    hoja.getRange(filaActual, 1).setFontWeight('bold').setBackground('#e3f2fd');
    filaActual++;

    var statsEmp = { horas: 0, pagar: 0, terapia: 0 };
    var currentIngreso = null;
    var lastEgreso = null;
    var fechaAnterior = null;

    for (var i = 0; i < registros.length; i++) {
      var reg = registros[i];

      // --- CORRECCIÓN: RESETEAR SI CAMBIA EL DÍA ---
      if (fechaAnterior && !esMismaFecha(fechaAnterior, reg.fecha)) {
        if (currentIngreso) {
          // SOLO ESTIMAR SI EL INGRESO FUE ANTES DE LAS 17:00 (5 PM)
          if (currentIngreso.fecha.getHours() < 17) {
             var salidaEst = new Date(currentIngreso.fecha.getTime() + (HORAS_JORNADA_NORMAL * 60 * 60 * 1000));
             var porc = 100;
             var tipoLabel = 'Normal (Estimado)';
             if (esDiaDeEstudio(empleadoId, currentIngreso.fecha, diasEstudioMapa)) {
               tipoLabel = 'Día de Estudio (Est.)';
               porc = 0;
             }
             escribirFilaReporte(hoja, filaActual++, currentIngreso.fecha, currentIngreso.fecha, salidaEst, tipoLabel + '*', HORAS_JORNADA_NORMAL, porc);
             statsEmp.horas += HORAS_JORNADA_NORMAL;
             statsEmp.pagar += HORAS_JORNADA_NORMAL * (porc / 100);
          }
          currentIngreso = null;
        }
        lastEgreso = null; 
      }
      fechaAnterior = reg.fecha;

      if (reg.esIngreso) {
        // REGLA: Ignorar entradas tarde (>= 5 PM) si es el primer registro o regreso no justificado
        var hora = reg.fecha.getHours();
        if (currentIngreso === null && lastEgreso === null && hora >= 17) continue;
        
        // Bloquear si ya hay un ingreso abierto el mismo día (evitar dobles marcas tarde)
        if (currentIngreso && esMismaFecha(currentIngreso.fecha, reg.fecha)) continue;

        if (lastEgreso && esMismaFecha(lastEgreso.fecha, reg.fecha)) {
          var enListaTerapia = listaTerapias[empleadoId] === true;
          if (enListaTerapia || lastEgreso.esTerapia) {
            var horasTerapia = (reg.fecha - lastEgreso.fecha) / (1000 * 60 * 60);
            if (horasTerapia > 0 && horasTerapia < 4) {
              escribirFilaReporte(hoja, filaActual++, reg.fecha, lastEgreso.fecha, reg.fecha, 'Terapia (Gap)', horasTerapia, 100);
              statsEmp.terapia += horasTerapia;
              statsEmp.pagar += horasTerapia;
            }
          }
        }
        currentIngreso = reg;
        lastEgreso = null;
      } 
      else if (reg.esEgreso) {
        if (currentIngreso) {
          var horas = (reg.fecha - currentIngreso.fecha) / (1000 * 60 * 60);
          var porc = reg.esPermiso ? 0 : (reg.esComputacion ? 50 : 100);
          var tipoLabel = reg.esPermiso ? 'Permiso' : (reg.esTerapia ? 'Terapia' : (reg.esComputacion ? 'Computación' : 'Normal'));
          
          if (esDiaDeEstudio(empleadoId, currentIngreso.fecha, diasEstudioMapa)) {
            tipoLabel = 'Día de Estudio';
            porc = 0;
          }

          escribirFilaReporte(hoja, filaActual++, currentIngreso.fecha, currentIngreso.fecha, reg.fecha, tipoLabel, horas, porc);
          
          var hPagar = horas * (porc / 100);
          statsEmp.horas += horas;
          statsEmp.pagar += hPagar;
          if (tipoLabel.indexOf('Terapia') !== -1) statsEmp.terapia += horas;
          
          currentIngreso = null;
          lastEgreso = reg;
        }
      }
    }

    if (currentIngreso && currentIngreso.fecha.getHours() < 17) {
        var salidaEst = new Date(currentIngreso.fecha.getTime() + (HORAS_JORNADA_NORMAL * 60 * 60 * 1000));
        var porc = 100;
        var tipoLabel = 'Normal (Estimado)';
        if (esDiaDeEstudio(empleadoId, currentIngreso.fecha, diasEstudioMapa)) {
          tipoLabel = 'Día de Estudio (Est.)';
          porc = 0;
        }
        escribirFilaReporte(hoja, filaActual++, currentIngreso.fecha, currentIngreso.fecha, salidaEst, tipoLabel + '*', HORAS_JORNADA_NORMAL, porc);
        statsEmp.horas += HORAS_JORNADA_NORMAL;
        statsEmp.pagar += HORAS_JORNADA_NORMAL * (porc / 100);
    }

    hoja.getRange(filaActual, 1, 1, 5).merge().setValue('SUBTOTAL ' + empleadoId.toUpperCase()).setFontWeight('bold').setHorizontalAlignment('right').setBackground('#f5f5f5');
    hoja.getRange(filaActual, 6).setValue(statsEmp.horas.toFixed(2)).setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment('center');
    hoja.getRange(filaActual, 8).setValue(statsEmp.pagar.toFixed(2)).setFontWeight('bold').setBackground('#f5f5f5').setHorizontalAlignment('center');
    filaActual++;
    if (statsEmp.terapia > 0) {
      hoja.getRange(filaActual, 1, 1, 8).merge().setValue('🧘 Total Terapia: ' + statsEmp.terapia.toFixed(2) + ' hrs').setFontStyle('italic').setFontSize(9).setFontColor('#00796b');
      filaActual++;
    }
    filaActual++;

    totalesGen.horas += statsEmp.horas;
    totalesGen.pagar += statsEmp.pagar;
    totalesGen.terapia += statsEmp.terapia;
  });

  // 3. Resumen Final
  filaActual++;
  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('RESUMEN GENERAL').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setBackground('#cfd8dc');
  filaActual++;
  
  var tablaResumen = [
    ['Total Personas:', listaEmpleados.length, 'Total Horas Lab:', totalesGen.horas.toFixed(2)],
    ['Total Terapias:', totalesGen.terapia.toFixed(2), 'TOTAL A PAGAR:', totalesGen.pagar.toFixed(2)]
  ];
  hoja.getRange(filaActual, 1, 2, 4).setValues(tablaResumen);
  hoja.getRange(filaActual, 3, 2, 1).setFontWeight('bold');
  hoja.getRange(filaActual + 1, 3, 1, 2).setBackground('#fff9c4').setFontWeight('bold');
  filaActual += 4;

  hoja.getRange(filaActual, 1, 1, 8).merge().setValue('⚠️ NOTA: * = Salida estimada (7 hrs). 🟢 = Entrada | 🔴 = Salida | 🧘 = Terapia | 💻 = Computación | 📚 = Estudio');
  hoja.getRange(filaActual, 1).setFontSize(8).setFontStyle('italic').setFontColor('#d32f2f');
  
  hoja.setColumnWidths(1, 8, 100);
  hoja.setColumnWidth(5, 130);
  hoja.activate();
  
  SpreadsheetApp.getUi().alert('✅ Reporte generado: ' + nombreHoja + '\nParticipantes: ' + listaEmpleados.length + '\nHoras totales: ' + totalesGen.pagar.toFixed(2));
}

// Función auxiliar para escribir filas
function escribirFilaReporte(hoja, fila, fechaRef, start, end, tipo, horas, porc) {
  var tz = Session.getScriptTimeZone();
  var hEntrada = Utilities.formatDate(start, tz, 'HH:mm') + (start.getHours() < 12 ? ' AM' : ' PM');
  var hSalida = Utilities.formatDate(end, tz, 'HH:mm') + (end.getHours() < 12 ? ' AM' : ' PM');
  
  var values = [[
    Utilities.formatDate(fechaRef, tz, 'dd/MM/yyyy'),
    DIAS_SEMANA[fechaRef.getDay()],
    hEntrada,
    hSalida,
    tipo,
    horas.toFixed(2),
    porc + '%',
    (horas * porc / 100).toFixed(2)
  ]];
  
  hoja.getRange(fila, 1, 1, 8).setValues(values).setHorizontalAlignment('center').setFontSize(9);
  
  // Colores por tipo
  var color = '#ffffff';
  if (tipo.indexOf('Terapia') !== -1) color = '#e0f2f1';       // Verde claro
  else if (tipo.indexOf('Computación') !== -1) color = '#fff9c4'; // Amarillo claro
  else if (tipo.indexOf('Permiso') !== -1 || porc === 0) color = '#ffebee'; // Rojo claro
  else if (tipo.indexOf('Estudio') !== -1) color = '#f3e5f5';   // Morado claro
  hoja.getRange(fila, 1, 1, 8).setBackground(color);
}

// ==================== DETECCIÓN DE COLUMNAS Y TIPO DE REGISTRO ====================

// Palabras clave que indican una acción de entrada/salida
var PALABRAS_ACCION = ['ingreso', 'egreso', 'entrada', 'salida', 'terapia', 'permiso', 'comput'];

// Verifica si un texto contiene alguna palabra de acción
function contieneAccion(texto) {
  var t = String(texto || '').toLowerCase();
  for (var i = 0; i < PALABRAS_ACCION.length; i++) {
    if (t.indexOf(PALABRAS_ACCION[i]) !== -1) return true;
  }
  return false;
}

// Detecta columnas del CSV de Kobo (soporta CUALQUIER formato)
// Recibe encabezados y opcionalmente las primeras filas de datos para auto-detectar
function detectarColumnas(encabezados, datosEjemplo) {
  var cols = {};

  // PASO 1: Detectar por nombre de encabezado
  for (var i = 0; i < encabezados.length; i++) {
    var h = String(encabezados[i]).trim();
    var hLow = h.toLowerCase();

    // Columnas estándar
    if (hLow === 'start') { cols.start = i; continue; }
    if (hLow === 'end') { cols.end = i; continue; }
    // UUID: preferir "_uuid" sobre "meta/rootUuid"
    if (hLow === '_uuid') { cols.uuid = i; continue; }
    if (hLow.indexOf('uuid') !== -1 && cols.uuid === undefined) { cols.uuid = i; continue; }

    // Detectar columnas de participante (puede haber varias: "Participante" y "Seleccione el nombre:")
    if (hLow === 'participante' || hLow.indexOf('participante') !== -1) {
      if (cols.participante === undefined) cols.participante = i;
      else if (cols.participante2 === undefined) cols.participante2 = i;
      continue;
    }
    if (hLow.indexOf('nombre') !== -1 || hLow.indexOf('seleccione') !== -1) {
      if (cols.participante === undefined) cols.participante = i;
      else if (cols.participante2 === undefined) cols.participante2 = i;
      continue;
    }

    // Columna de acción unificada (select_one)
    // Detectar si el encabezado contiene palabras de entrada Y salida
    var tieneEntrada = hLow.indexOf('ingreso') !== -1 || hLow.indexOf('entrada') !== -1;
    var tieneSalida = hLow.indexOf('egreso') !== -1 || hLow.indexOf('salida') !== -1;
    if (tieneEntrada && tieneSalida) { cols.accionUnificada = i; continue; }

    // Detectar "accion", "acción", "type", "marcar", "registro"
    if (hLow.indexOf('accion') !== -1 || hLow.indexOf('acción') !== -1 ||
        hLow === 'type' || hLow.indexOf('marcar') !== -1 || hLow.indexOf('registro') !== -1) {
      cols.accionUnificada = i;
      continue;
    }

    // Columna de subtipo (agregada por normalización)
    if (hLow === 'subtipo_egreso' || (hLow.indexOf('subtipo') !== -1 && hLow.indexOf('egreso') !== -1)) { cols.subtipoEgreso = i; continue; }

    // FORMATO VIEJO (select_multiple): columnas separadas
    if (hLow.indexOf('/ingreso') !== -1 || hLow.indexOf('/entrada') !== -1) { cols.ingreso = i; continue; }
    if (hLow.indexOf('/egreso') !== -1 || hLow.indexOf('/salida') !== -1) { cols.egreso = i; continue; }
    if (hLow.indexOf('/terapia') !== -1) { cols.terapia = i; continue; }
    if (hLow.indexOf('/permiso') !== -1) { cols.permiso = i; continue; }
    if (hLow.indexOf('/comput') !== -1) { cols.computacion = i; continue; }
  }

  // PASO 2: Si no encontró columna de acción, ESCANEAR los datos para auto-detectar
  if (cols.accionUnificada === undefined && cols.ingreso === undefined && cols.egreso === undefined && datosEjemplo) {
    // Contar cuántas celdas de cada columna contienen palabras de acción
    var puntuacion = [];
    for (var c = 0; c < encabezados.length; c++) {
      // Saltar columnas ya identificadas (start, end, participante, uuid)
      if (c === cols.start || c === cols.end || c === cols.participante || c === cols.uuid) {
        puntuacion.push(0);
        continue;
      }
      var aciertos = 0;
      var maxFilas = Math.min(datosEjemplo.length, 50); // Escanear máximo 50 filas
      for (var f = 0; f < maxFilas; f++) {
        if (contieneAccion(datosEjemplo[f][c])) aciertos++;
      }
      puntuacion.push(aciertos);
    }

    // La columna con más aciertos es la de acción
    var mejorCol = -1;
    var mejorPunt = 0;
    for (var c = 0; c < puntuacion.length; c++) {
      if (puntuacion[c] > mejorPunt) {
        mejorPunt = puntuacion[c];
        mejorCol = c;
      }
    }

    // Si al menos 30% de las filas tienen una palabra de acción, es la columna correcta
    if (mejorCol !== -1 && mejorPunt >= Math.max(1, datosEjemplo.length * 0.3)) {
      cols.accionUnificada = mejorCol;
      cols.accionAutoDetectada = true; // Flag para el diagnóstico
    }
  }

  // PASO 3: Si tampoco encontró "start", buscar columna con fechas ISO
  if (cols.start === undefined && datosEjemplo) {
    for (var c = 0; c < encabezados.length; c++) {
      if (c === cols.participante || c === cols.accionUnificada) continue;
      var val = String(datosEjemplo[0] ? datosEjemplo[0][c] : '');
      // Detectar formato ISO: 2026-03-27T16:10:01
      if (val.match(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/)) {
        cols.start = c;
        break;
      }
    }
  }

  return cols;
}

// Busca las columnas de participante en un array de encabezados
// Devuelve { col1: indice, col2: indice } o { col1: -1 }
function buscarColumnasParticipante(encabezados) {
  var result = { col1: -1, col2: -1 };
  for (var i = 0; i < encabezados.length; i++) {
    var h = String(encabezados[i]).trim().toLowerCase();
    if (h.indexOf('participante') !== -1 || h.indexOf('nombre') !== -1 || h.indexOf('seleccione') !== -1) {
      if (result.col1 === -1) result.col1 = i;
      else if (result.col2 === -1) result.col2 = i;
    }
  }
  return result;
}

// Obtiene el nombre del participante de una fila, dadas las columnas encontradas
// Versión simple para funciones que no usan detectarColumnas
function obtenerNombreDeFila(row, colsParticipante) {
  var nombre = colsParticipante.col1 !== -1 ? String(row[colsParticipante.col1] || '').trim() : '';
  if (!nombre && colsParticipante.col2 !== -1) {
    nombre = String(row[colsParticipante.col2] || '').trim();
  }
  if (nombre && colsParticipante.col2 !== -1) {
    var nombre2 = String(row[colsParticipante.col2] || '').trim();
    if (nombre2 && nombre2.length > nombre.length) nombre = nombre2;
  }
  return nombre;
}

// Obtiene el nombre del participante de una fila, probando ambas columnas
function obtenerParticipanteFila(row, cols) {
  // Probar columna principal primero
  var nombre = cols.participante !== undefined ? String(row[cols.participante] || '').trim() : '';
  // Si está vacía, probar columna secundaria
  if (!nombre && cols.participante2 !== undefined) {
    nombre = String(row[cols.participante2] || '').trim();
  }
  // Si la principal tiene datos pero la secundaria tiene un nombre más completo, preferir la secundaria
  if (nombre && cols.participante2 !== undefined) {
    var nombre2 = String(row[cols.participante2] || '').trim();
    if (nombre2 && nombre2.length > nombre.length) nombre = nombre2;
  }
  return nombre;
}

// Determina el tipo de un registro (fila) según el formato del CSV
function obtenerTipoRegistro(row, cols) {
  var resultado = { tipo: '', esIngreso: false, esEgreso: false, esTerapia: false, esPermiso: false, esComputacion: false };

  if (cols.accionUnificada !== undefined) {
    // FORMATO NUEVO: leer el valor de la celda
    var valOriginal = String(row[cols.accionUnificada] || '').trim();
    var val = valOriginal.toLowerCase();
    resultado.tipo = val;

    // Detectar tipo — valores canónicos: "Entrada" / "Salida"
    // También reconoce variantes antiguas: Ingreso, Egreso, 🟢 ENTRADA, 🔴 SALIDA
    if (val.indexOf('entrada') !== -1 || val.indexOf('ingreso') !== -1 ||
        valOriginal.indexOf('🟢') !== -1) {
      resultado.esIngreso = true;
    }
    if (val.indexOf('salida') !== -1 || val.indexOf('egreso') !== -1 ||
        valOriginal.indexOf('🔴') !== -1) {
      resultado.esEgreso = true;
    }

    // Si tiene ambos (raro pero posible), priorizar lo que dice
    // Si no tiene ni ingreso ni egreso, revisar otros tipos
    if (!resultado.esIngreso && !resultado.esEgreso) {
      // Terapia como tipo separado = es una salida para terapia
      if (val.indexOf('terapia') !== -1) {
        resultado.esTerapia = true;
        resultado.esEgreso = true;
      }
      // Permiso = salida con permiso
      if (val.indexOf('permiso') !== -1) {
        resultado.esPermiso = true;
        resultado.esEgreso = true;
      }
      // Computación = salida para computación
      if (val.indexOf('comput') !== -1) {
        resultado.esComputacion = true;
        resultado.esEgreso = true;
      }
    } else {
      // Si ya es ingreso o egreso, verificar si TAMBIÉN es terapia/permiso/computación
      if (val.indexOf('terapia') !== -1) resultado.esTerapia = true;
      if (val.indexOf('permiso') !== -1) resultado.esPermiso = true;
      if (val.indexOf('comput') !== -1) resultado.esComputacion = true;
    }

    // Leer subtipo desde columna subtipo_egreso (preserva subtipo aunque la acción sea solo "Salida")
    if (resultado.esEgreso && cols.subtipoEgreso !== undefined && row[cols.subtipoEgreso]) {
      var sub = String(row[cols.subtipoEgreso]).trim().toLowerCase();
      if (sub.indexOf('terapia') !== -1) resultado.esTerapia = true;
      if (sub.indexOf('permiso') !== -1) resultado.esPermiso = true;
      if (sub.indexOf('comput') !== -1) resultado.esComputacion = true;
    }
  } else {
    // FORMATO VIEJO: columnas separadas con TRUE/FALSE
    if (cols.ingreso !== undefined && row[cols.ingreso]) {
      var v = String(row[cols.ingreso]).trim().toLowerCase();
      if (v === 'true' || v === '1' || v === 'x' || v === 'yes') {
        resultado.esIngreso = true;
        resultado.tipo = 'entrada';
      }
    }
    if (cols.egreso !== undefined && row[cols.egreso]) {
      var v = String(row[cols.egreso]).trim().toLowerCase();
      if (v === 'true' || v === '1' || v === 'x' || v === 'yes') {
        resultado.esEgreso = true;
        resultado.tipo = 'salida';
      }
    }
    if (cols.terapia !== undefined && row[cols.terapia]) {
      var v = String(row[cols.terapia]).trim().toLowerCase();
      if (v === 'true' || v === '1' || v === 'x' || v === 'yes') {
        resultado.esTerapia = true;
        resultado.esEgreso = true;
        resultado.tipo = 'terapia';
      }
    }
    if (cols.permiso !== undefined && row[cols.permiso]) {
      var v = String(row[cols.permiso]).trim().toLowerCase();
      if (v === 'true' || v === '1' || v === 'x' || v === 'yes') {
        resultado.esPermiso = true;
        resultado.esEgreso = true;
        resultado.tipo = 'permiso';
      }
    }
    if (cols.computacion !== undefined && row[cols.computacion]) {
      var v = String(row[cols.computacion]).trim().toLowerCase();
      if (v === 'true' || v === '1' || v === 'x' || v === 'yes') {
        resultado.esComputacion = true;
        resultado.esEgreso = true;
        resultado.tipo = 'computación';
      }
    }
  }

  return resultado;
}

// ==================== NORMALIZAR TODO ====================
// Ejecuta en secuencia las 4 fases de normalización:
// Fase 1: Repara typos en Ingreso/Egreso (DatosKobo)
// Fase 2: Crea/actualiza NombresCanonicos
// Fase 3: Normaliza nombres en DatosKobo
// Fase 4: Actualiza DiasEstudio y ListaTerapias
function normalizarTodo() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = spreadsheet.getSheetByName("DatosKobo");

  if (!hojaKobo) {
    ui.alert('❌ No existe DatosKobo.\nPrimero importa datos desde Kobo con "🔄 Actualizar desde Kobo".');
    return;
  }

  var resp = ui.alert(
    '✨ NORMALIZAR TODO',
    'Se ejecutarán 4 fases automáticamente:\n\n' +
    '1️⃣  Normalizar Entrada/Salida (ej: "Ingreso" → "Entrada", "Egreso" → "Salida")\n' +
    '2️⃣  Crear/actualizar tabla de nombres canónicos\n' +
    '3️⃣  Normalizar nombres en DatosKobo\n' +
    '4️⃣  Actualizar DiasEstudio y ListaTerapias\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var log = [];
  var datos = hojaKobo.getDataRange().getValues();
  var encabezados = datos[0];
  var cols = detectarColumnas(encabezados, datos.slice(1));

  // ── FASE 1: Reparar typos en Ingreso/Egreso ──────────────────────
  SpreadsheetApp.getActiveSpreadsheet().toast('Fase 1/4: Reparando Ingreso/Egreso...', '✨', -1);

  // Terapia, Permiso y Computacion son subcategorías de Salida → se normalizan a '🔴 Salida'
  var CONOCIDOS = [
    { clave: 'ingreso', correcto: '🟢 Entrada' },
    { clave: 'entrada', correcto: '🟢 Entrada' },
    { clave: 'egreso',  correcto: '🔴 Salida'  },
    { clave: 'salida',  correcto: '🔴 Salida'  },
    { clave: 'terapia', correcto: '🔴 Salida'  },
    { clave: 'permiso', correcto: '🔴 Salida'  },
    { clave: 'comput',  correcto: '🔴 Salida'  }
  ];

  var cambiosAccion = 0;
  var sinCorreccion = [];

  // Mapa de subtipo: qué palabra clave indica qué subtipo canónico
  var SUBTIPOS_MAP = [
    { clave: 'terapia', subtipo: 'Terapia' },
    { clave: 'permiso', subtipo: 'Permiso' },
    { clave: 'comput',  subtipo: 'Computacion' }
  ];

  if (cols.accionUnificada !== undefined) {
    // Asegurar que existe la columna subtipo_egreso en DatosKobo
    var colSubtipo = cols.subtipoEgreso;
    if (colSubtipo === undefined) {
      var numCols = encabezados.length;
      hojaKobo.getRange(1, numCols + 1).setValue('subtipo_egreso');
      hojaKobo.getRange(1, numCols + 1).setFontWeight('bold').setBackground('#e6b8a2').setFontColor('#000000');
      colSubtipo = numCols; // índice 0-based
      cols.subtipoEgreso = colSubtipo;
    }

    for (var f = 1; f < datos.length; f++) {
      var valRaw = String(datos[f][cols.accionUnificada] || '').trim();
      if (!valRaw) continue;

      var tipo = obtenerTipoRegistro(datos[f], cols);
      var correcto = null;
      var subtipo = '';

      // Detectar subtipo desde el valor original ANTES de normalizar
      var valLowOrig = valRaw.toLowerCase();
      for (var s = 0; s < SUBTIPOS_MAP.length; s++) {
        if (valLowOrig.indexOf(SUBTIPOS_MAP[s].clave) !== -1) {
          subtipo = SUBTIPOS_MAP[s].subtipo;
          break;
        }
      }

      // Estandarizar a forma canónica
      if (tipo.esIngreso) {
        correcto = '🟢 Entrada';
        subtipo = ''; // Las entradas no tienen subtipo
      } else if (tipo.esEgreso) {
        correcto = '🔴 Salida';
      } else {
        // Typo: intentar corrección por similitud
        var valLow = valRaw.toLowerCase();
        for (var k = 0; k < CONOCIDOS.length; k++) {
          var clave = CONOCIDOS[k].clave;
          var shared = 0;
          for (var c = 0; c < Math.min(valLow.length, clave.length); c++) {
            if (valLow[c] === clave[c]) shared++;
          }
          var sim = shared / Math.max(valLow.length, clave.length);
          var contiene = valLow.indexOf(clave.substring(0, 4)) !== -1 || clave.indexOf(valLow.substring(0, 4)) !== -1;
          if (sim >= 0.5 || contiene) {
            correcto = CONOCIDOS[k].correcto;
            // Detectar subtipo también para typos
            for (var s2 = 0; s2 < SUBTIPOS_MAP.length; s2++) {
              if (clave.indexOf(SUBTIPOS_MAP[s2].clave) !== -1) {
                subtipo = SUBTIPOS_MAP[s2].subtipo;
                break;
              }
            }
            break;
          }
        }
      }

      var celdaAccion = hojaKobo.getRange(f + 1, cols.accionUnificada + 1);
      if (correcto && valRaw !== correcto) {
        celdaAccion.setValue(correcto);
        cambiosAccion++;
      }
      // Colorear verde=🟢 Entrada, rojo=🔴 Salida (siempre, aunque el valor ya fuera correcto)
      if (correcto === '🟢 Entrada') {
        celdaAccion.setBackground('#b7e1cd').setFontColor('#0b5c30').setFontWeight('bold');
      } else if (correcto === '🔴 Salida') {
        celdaAccion.setBackground('#f4cccc').setFontColor('#7f0000').setFontWeight('bold');
      }

      if (!correcto && sinCorreccion.indexOf(valRaw) === -1) {
        sinCorreccion.push(valRaw);
      }

      // Escribir subtipo si aplica y la celda aún está vacía (no sobreescribir lo que el usuario puso)
      var subtipoActual = String(datos[f][colSubtipo] || '').trim();
      if (subtipo && !subtipoActual) {
        hojaKobo.getRange(f + 1, colSubtipo + 1).setValue(subtipo);
      }
    }
  }
  // Formatear columnas start/end → fecha legible dd/MM/yyyy HH:mm
  normalizarAccionSilencioso(hojaKobo);

  log.push('1️⃣  Entrada/Salida normalizados: ' + cambiosAccion + ' registros' +
    (sinCorreccion.length > 0 ? ' (⚠️ sin corrección: "' + sinCorreccion.join('", "') + '")' : ''));

  // ── FASE 2: Crear/actualizar NombresCanonicos ─────────────────────
  SpreadsheetApp.getActiveSpreadsheet().toast('Fase 2/4: Actualizando NombresCanonicos...', '✨', -1);

  // Recargar datos (pueden haber cambiado en fase 1)
  datos = hojaKobo.getDataRange().getValues();
  var colsP = buscarColumnasParticipante(datos[0]);
  var nombresUnicos = {};
  for (var f = 1; f < datos.length; f++) {
    var nombre = obtenerNombreDeFila(datos[f], colsP);
    if (nombre) nombresUnicos[nombre] = true;
  }
  var todosNombres = Object.keys(nombresUnicos);

  // Agrupar por código
  var porCodigo = {};
  var sinCodigo = [];
  for (var i = 0; i < todosNombres.length; i++) {
    var n = todosNombres[i];
    var codigo = extraerCodigo(n);
    if (codigo) {
      if (!porCodigo[codigo]) porCodigo[codigo] = [];
      porCodigo[codigo].push(n);
    } else {
      sinCodigo.push(n);
    }
  }
  for (var i = 0; i < sinCodigo.length; i++) {
    var nombreSC = sinCodigo[i];
    var limSC = textoParaComparar(limpiarNombre(nombreSC));
    var encontrado = false;
    var codsKeys = Object.keys(porCodigo);
    for (var c = 0; c < codsKeys.length; c++) {
      var grupo = porCodigo[codsKeys[c]];
      for (var g = 0; g < grupo.length; g++) {
        if (nombresCoinciden(limSC, textoParaComparar(limpiarNombre(grupo[g])))) {
          grupo.push(nombreSC);
          encontrado = true;
          break;
        }
      }
      if (encontrado) break;
    }
    if (!encontrado) porCodigo['_SIN_' + i] = [nombreSC];
  }

  // Elegir nombre canónico por grupo
  var mapeo = {};
  var codsKeys2 = Object.keys(porCodigo);
  for (var c = 0; c < codsKeys2.length; c++) {
    var grupo = porCodigo[codsKeys2[c]];
    var codigoReal = codsKeys2[c].indexOf('_SIN_') === 0 ? '' : codsKeys2[c];
    var mejorNombre = '';
    var mejorLargo = 0;
    for (var g = 0; g < grupo.length; g++) {
      var lim = limpiarNombre(grupo[g]);
      if (lim.length > mejorLargo) { mejorLargo = lim.length; mejorNombre = lim; }
    }
    // Formato: "Nombre (CÓDIGO)" — código SIEMPRE al final entre paréntesis
    var canonico = codigoReal ? mejorNombre + ' (' + codigoReal + ')' : mejorNombre;
    for (var g = 0; g < grupo.length; g++) mapeo[grupo[g]] = canonico;
  }

  // Preservar correcciones manuales del usuario
  var hojaNombres = spreadsheet.getSheetByName("NombresCanonicos");
  if (hojaNombres) {
    var datosExist = hojaNombres.getDataRange().getValues();
    for (var f = 1; f < datosExist.length; f++) {
      var orig = String(datosExist[f][0] || '').trim();
      var canon = String(datosExist[f][1] || '').trim();
      if (orig && canon && mapeo[orig] !== undefined) mapeo[orig] = canon;
    }
  }

  // Escribir hoja NombresCanonicos
  if (!hojaNombres) hojaNombres = spreadsheet.insertSheet("NombresCanonicos");
  else hojaNombres.clearContents();

  hojaNombres.getRange(1, 1, 1, 3).setValues([['Nombre Original (Kobo)', 'Nombre Canónico', 'Código']]);
  hojaNombres.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#ff6f00').setFontColor('#ffffff').setHorizontalAlignment('center');
  hojaNombres.setFrozenRows(1);

  var filas = [];
  var claves = Object.keys(mapeo).sort();
  var duplicados = 0;
  for (var i = 0; i < claves.length; i++) {
    var cod = extraerCodigo(claves[i]) || '';
    filas.push([claves[i], mapeo[claves[i]], cod]);
    if (claves[i] !== mapeo[claves[i]]) duplicados++;
  }
  if (filas.length > 0) {
    hojaNombres.getRange(2, 1, filas.length, 3).setValues(filas);
    for (var i = 0; i < filas.length; i++) {
      if (filas[i][0] !== filas[i][1]) hojaNombres.getRange(i + 2, 1, 1, 3).setBackground('#fff3e0');
    }
  }
  hojaNombres.setColumnWidth(1, 350);
  hojaNombres.setColumnWidth(2, 350);
  hojaNombres.setColumnWidth(3, 130);

  log.push('2️⃣  NombresCanonicos: ' + claves.length + ' nombres, ' + duplicados + ' variantes agrupadas');

  // ── FASE 3: Normalizar nombres en DatosKobo ───────────────────────
  SpreadsheetApp.getActiveSpreadsheet().toast('Fase 3/4: Normalizando nombres en DatosKobo...', '✨', -1);

  datos = hojaKobo.getDataRange().getValues();
  cols = detectarColumnas(datos[0], datos.slice(1));
  var cambiosNombreKobo = 0;

  for (var f = 1; f < datos.length; f++) {
    if (cols.participante !== undefined) {
      var n1 = String(datos[f][cols.participante] || '').trim();
      if (n1 && mapeo[n1] && mapeo[n1] !== n1) {
        hojaKobo.getRange(f + 1, cols.participante + 1).setValue(mapeo[n1]);
        cambiosNombreKobo++;
      }
    }
    if (cols.participante2 !== undefined) {
      var n2 = String(datos[f][cols.participante2] || '').trim();
      if (n2 && mapeo[n2] && mapeo[n2] !== n2) {
        hojaKobo.getRange(f + 1, cols.participante2 + 1).setValue(mapeo[n2]);
        cambiosNombreKobo++;
      }
    }
  }
  log.push('3️⃣  DatosKobo nombres normalizados: ' + cambiosNombreKobo + ' celdas');

  // ── FASE 4: Actualizar DiasEstudio y ListaTerapias ────────────────
  SpreadsheetApp.getActiveSpreadsheet().toast('Fase 4/4: Actualizando DiasEstudio y ListaTerapias...', '✨', -1);

  var hojasActualizadas = normalizarNombresEnHojas(mapeo, spreadsheet);
  log.push('4️⃣  Hojas actualizadas: ' + (hojasActualizadas || 'ninguna pendiente'));

  // ── RESULTADO ─────────────────────────────────────────────────────
  SpreadsheetApp.getActiveSpreadsheet().toast('', '', 1);
  var resumen = '✅ NORMALIZACIÓN COMPLETA\n\n' + log.join('\n');
  if (sinCorreccion.length > 0) {
    resumen += '\n\n⚠️ Estos valores no pudieron corregirse automáticamente:\n"' +
      sinCorreccion.join('"\n"') + '"\n' +
      'Edítalos manualmente en DatosKobo (columna "Ingreso / Egreso").';
  }
  resumen += '\n\nYa puedes generar el reporte.';
  ui.alert(resumen);
}

// ==================== REPARADOR DE DATOS KOBO ====================
// Corrige typos en Ingreso/Egreso y normaliza nombres en DatosKobo.
// Se ejecuta en 2 fases: primero diagnóstica, luego corrige con confirmación.
function repararDatosKobo() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = spreadsheet.getSheetByName("DatosKobo");

  if (!hojaKobo) {
    ui.alert('No existe DatosKobo. Importa primero desde Kobo.');
    return;
  }

  var datos = hojaKobo.getDataRange().getValues();
  var encabezados = datos[0];
  var cols = detectarColumnas(encabezados, datos.slice(1));

  // ---------- FASE 1: Detectar problemas ----------
  var valoresDesconocidos = {}; // val → count
  var nombresANormalizar = 0;
  var mapeoNombres = cargarMapeoNombres();

  // Escanear acción
  if (cols.accionUnificada !== undefined) {
    for (var f = 1; f < datos.length; f++) {
      var tipo = obtenerTipoRegistro(datos[f], cols);
      if (!tipo.esIngreso && !tipo.esEgreso) {
        var valRaw = String(datos[f][cols.accionUnificada] || '').trim();
        if (valRaw) {
          valoresDesconocidos[valRaw] = (valoresDesconocidos[valRaw] || 0) + 1;
        }
      }
    }
  }

  // Escanear nombres
  if (Object.keys(mapeoNombres).length > 0 && cols.participante !== undefined) {
    for (var f = 1; f < datos.length; f++) {
      var n1 = String(datos[f][cols.participante] || '').trim();
      if (n1 && mapeoNombres[n1] && mapeoNombres[n1] !== n1) nombresANormalizar++;
    }
  }

  var listaDesc = Object.keys(valoresDesconocidos);
  var totalDesc = listaDesc.reduce(function(s, k) { return s + valoresDesconocidos[k]; }, 0);

  // Proponer correcciones por similitud
  var CONOCIDOS = [
    { clave: 'ingreso', correcto: '🟢 Entrada' },
    { clave: 'entrada', correcto: '🟢 Entrada' },
    { clave: 'egreso',  correcto: '🔴 Salida'  },
    { clave: 'salida',  correcto: '🔴 Salida'  },
    { clave: 'terapia', correcto: '🔴 Salida'  },
    { clave: 'permiso', correcto: '🔴 Salida'  },
    { clave: 'comput',  correcto: '🔴 Salida'  }
  ];

  var mapeoCorrecciones = {}; // valor desconocido → corrección propuesta
  for (var i = 0; i < listaDesc.length; i++) {
    var v = listaDesc[i].toLowerCase();
    for (var k = 0; k < CONOCIDOS.length; k++) {
      var clave = CONOCIDOS[k].clave;
      // Coincide si comparten los primeros 4 caracteres o uno contiene al otro
      var shared = 0;
      for (var c = 0; c < Math.min(v.length, clave.length); c++) {
        if (v[c] === clave[c]) shared++;
      }
      var sim = shared / Math.max(v.length, clave.length);
      // También verificar si uno contiene al otro
      var contiene = v.indexOf(clave.substring(0, 4)) !== -1 || clave.indexOf(v.substring(0, 4)) !== -1;
      if (sim >= 0.5 || contiene) {
        mapeoCorrecciones[listaDesc[i]] = CONOCIDOS[k].correcto;
        break;
      }
    }
  }

  // Construir mensaje de diagnóstico
  var msg = '🔧 REPARAR DATOS KOBO\n\n';

  if (listaDesc.length === 0) {
    msg += '✅ Sin problemas en Entrada/Salida.\n';
  } else {
    msg += '⚠️ FASE 1 - Valores no reconocidos (' + totalDesc + ' registros):\n';
    for (var i = 0; i < listaDesc.length; i++) {
      var prop = mapeoCorrecciones[listaDesc[i]] ? ' → "' + mapeoCorrecciones[listaDesc[i]] + '"' : ' → ❓ sin corrección';
      msg += '• "' + listaDesc[i] + '" (' + valoresDesconocidos[listaDesc[i]] + ' veces)' + prop + '\n';
    }
  }

  if (nombresANormalizar > 0) {
    msg += '\n📋 FASE 2 - Nombres a normalizar en DatosKobo: ' + nombresANormalizar + ' celdas\n';
  } else if (Object.keys(mapeoNombres).length === 0) {
    msg += '\n📋 FASE 2 - No hay hoja NombresCanonicos (ejecuta "Normalizar Nombres" primero).\n';
  } else {
    msg += '\n✅ Todos los nombres ya están normalizados en DatosKobo.\n';
  }

  if (listaDesc.length === 0 && nombresANormalizar === 0) {
    ui.alert(msg + '\nNo hay nada que reparar.');
    return;
  }

  msg += '\n¿Aplicar correcciones automáticas?';
  var resp = ui.alert('🔧 REPARAR DATOS', msg, ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // ---------- FASE 1: Corregir typos en acción ----------
  var cambiosAccion = 0;
  var sinCorreccion = [];
  if (cols.accionUnificada !== undefined && listaDesc.length > 0) {
    for (var f = 1; f < datos.length; f++) {
      var tipo = obtenerTipoRegistro(datos[f], cols);
      if (!tipo.esIngreso && !tipo.esEgreso) {
        var valRaw = String(datos[f][cols.accionUnificada] || '').trim();
        if (valRaw && mapeoCorrecciones[valRaw]) {
          hojaKobo.getRange(f + 1, cols.accionUnificada + 1).setValue(mapeoCorrecciones[valRaw]);
          cambiosAccion++;
        } else if (valRaw && !mapeoCorrecciones[valRaw]) {
          if (sinCorreccion.indexOf(valRaw) === -1) sinCorreccion.push(valRaw);
        }
      }
    }
  }

  // ---------- FASE 2: Normalizar nombres en DatosKobo ----------
  var cambiosNombre = 0;
  if (Object.keys(mapeoNombres).length > 0 && cols.participante !== undefined) {
    // Recargar datos (por si cambió algo en fase 1)
    var datosActuales = hojaKobo.getDataRange().getValues();
    for (var f = 1; f < datosActuales.length; f++) {
      var n1 = String(datosActuales[f][cols.participante] || '').trim();
      if (n1 && mapeoNombres[n1] && mapeoNombres[n1] !== n1) {
        hojaKobo.getRange(f + 1, cols.participante + 1).setValue(mapeoNombres[n1]);
        cambiosNombre++;
      }
      if (cols.participante2 !== undefined) {
        var n2 = String(datosActuales[f][cols.participante2] || '').trim();
        if (n2 && mapeoNombres[n2] && mapeoNombres[n2] !== n2) {
          hojaKobo.getRange(f + 1, cols.participante2 + 1).setValue(mapeoNombres[n2]);
          cambiosNombre++;
        }
      }
    }
  }

  // Resultado final
  var resumen = '✅ REPARACIÓN COMPLETADA\n\n';
  resumen += 'Fase 1 - Ingreso/Egreso corregidos: ' + cambiosAccion + ' registros\n';
  resumen += 'Fase 2 - Nombres normalizados: ' + cambiosNombre + ' celdas\n';
  if (sinCorreccion.length > 0) {
    resumen += '\n⚠️ Valores sin corrección automática (' + sinCorreccion.length + '):\n';
    for (var i = 0; i < sinCorreccion.length; i++) {
      resumen += '• "' + sinCorreccion[i] + '"\n';
    }
    resumen += '\nEdita estos directamente en la hoja DatosKobo (columna "Ingreso / Egreso").';
  }
  resumen += '\n\nYa puedes generar el reporte.';
  ui.alert(resumen);
}

// ==================== CAMBIAR NOMBRE DE PARTICIPANTE ====================
// Permite cambiar el nombre canónico de un participante en NombresCanonicos,
// DiasEstudio y ListaTerapias desde una sola función.
function cambiarNombreParticipante() {
  var ui = SpreadsheetApp.getUi();
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Verificar que exista NombresCanonicos
  var hojaNombres = spreadsheet.getSheetByName("NombresCanonicos");
  if (!hojaNombres) {
    ui.alert('⚠️ No existe la hoja "NombresCanonicos".\nPrimero ejecuta "🔀 Normalizar Nombres" desde el menú.');
    return;
  }

  // Leer nombres canónicos actuales (únicos)
  var datosNombres = hojaNombres.getDataRange().getValues();
  var canonicosUnicos = {};
  for (var f = 1; f < datosNombres.length; f++) {
    var canonico = String(datosNombres[f][1] || '').trim();
    if (canonico) canonicosUnicos[canonico] = true;
  }
  var listaCanonicos = Object.keys(canonicosUnicos).sort();

  // Mostrar lista y pedir nombre actual
  var listaTexto = listaCanonicos.slice(0, 25).join('\n');
  if (listaCanonicos.length > 25) listaTexto += '\n... (' + (listaCanonicos.length - 25) + ' más)';

  var resp1 = ui.prompt(
    '✏️ CAMBIAR NOMBRE DE PARTICIPANTE (1/2)',
    'Participantes actuales:\n' + listaTexto + '\n\nEscribe el nombre ACTUAL (exacto, copia y pega si es posible):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp1.getSelectedButton() !== ui.Button.OK) return;
  var nombreActual = resp1.getResponseText().trim();
  if (!nombreActual) return;

  // Verificar que existe
  if (!canonicosUnicos[nombreActual]) {
    // Buscar coincidencia parcial
    var sugerencias = listaCanonicos.filter(function(n) {
      return n.toLowerCase().indexOf(nombreActual.toLowerCase()) !== -1;
    });
    var msg = '⚠️ No se encontró "' + nombreActual + '" exactamente.';
    if (sugerencias.length > 0) {
      msg += '\n\n¿Quisiste decir alguno de estos?\n' + sugerencias.slice(0, 5).join('\n');
    }
    ui.alert(msg);
    return;
  }

  // Pedir nuevo nombre
  var resp2 = ui.prompt(
    '✏️ CAMBIAR NOMBRE DE PARTICIPANTE (2/2)',
    'Nombre actual:\n"' + nombreActual + '"\n\n¿Cuál será el NUEVO nombre?\n(Si tiene código, formato: "Nombre Completo (CÓDIGO)" )',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp2.getSelectedButton() !== ui.Button.OK) return;
  var nombreNuevo = resp2.getResponseText().trim();
  if (!nombreNuevo || nombreNuevo === nombreActual) return;

  // 1. Actualizar en NombresCanonicos (columna B = nombre canónico)
  var cambiados = 0;
  for (var f = 1; f < datosNombres.length; f++) {
    if (String(datosNombres[f][1] || '').trim() === nombreActual) {
      hojaNombres.getRange(f + 1, 2).setValue(nombreNuevo);
      cambiados++;
    }
  }

  // 2. Actualizar en DiasEstudio (columna A = nombre)
  var cambiosDias = 0;
  var hojaDias = spreadsheet.getSheetByName("DiasEstudio");
  if (hojaDias) {
    var datosDias = hojaDias.getDataRange().getValues();
    for (var f = 1; f < datosDias.length; f++) {
      if (String(datosDias[f][0] || '').trim() === nombreActual) {
        hojaDias.getRange(f + 1, 1).setValue(nombreNuevo);
        cambiosDias++;
      }
    }
  }

  // 3. Actualizar en ListaTerapias (columna A = nombre)
  var cambiosTer = 0;
  var hojaTer = spreadsheet.getSheetByName("ListaTerapias");
  if (hojaTer) {
    var datosTer = hojaTer.getDataRange().getValues();
    for (var f = 1; f < datosTer.length; f++) {
      if (String(datosTer[f][0] || '').trim() === nombreActual) {
        hojaTer.getRange(f + 1, 1).setValue(nombreNuevo);
        cambiosTer++;
      }
    }
  }

  var msg = '✅ NOMBRE CAMBIADO\n\n';
  msg += '"' + nombreActual + '"\n→ "' + nombreNuevo + '"\n\n';
  msg += 'NombresCanonicos: ' + cambiados + ' fila(s) actualizada(s)\n';
  if (cambiosDias > 0) msg += 'DiasEstudio: ' + cambiosDias + ' fila(s)\n';
  if (cambiosTer > 0) msg += 'ListaTerapias: ' + cambiosTer + ' fila(s)\n';
  msg += '\nLos próximos reportes usarán el nuevo nombre.';
  ui.alert(msg);
}

// Función de diagnóstico - muestra qué columnas detectó y ejemplos de datos
function diagnosticarDatosKobo() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var hojaKobo = spreadsheet.getSheetByName("DatosKobo");

  if (!hojaKobo) {
    SpreadsheetApp.getUi().alert('No existe la hoja "DatosKobo". Importa primero desde Kobo.');
    return;
  }

  var datos = hojaKobo.getDataRange().getValues();
  var encabezados = datos[0];

  // Mostrar todos los encabezados
  var listaEnc = '';
  for (var i = 0; i < encabezados.length; i++) {
    listaEnc += 'Col ' + i + ': "' + encabezados[i] + '"\n';
  }

  // Detectar columnas (con auto-detección por escaneo de datos)
  var datosFilas = datos.slice(1);
  var cols = detectarColumnas(encabezados, datosFilas);
  var deteccion = '\n--- COLUMNAS DETECTADAS ---\n';
  deteccion += 'start: ' + (cols.start !== undefined ? 'Col ' + cols.start : '❌ NO ENCONTRADA') + '\n';
  deteccion += 'end: ' + (cols.end !== undefined ? 'Col ' + cols.end : 'NO (opcional)') + '\n';
  deteccion += 'participante: ' + (cols.participante !== undefined ? 'Col ' + cols.participante + ' ("' + encabezados[cols.participante] + '")' : '❌ NO ENCONTRADA') + '\n';
  if (cols.participante2 !== undefined) {
    deteccion += 'participante2: Col ' + cols.participante2 + ' ("' + encabezados[cols.participante2] + '") [se usa el más completo]\n';
  }
  deteccion += 'acción: ' + (cols.accionUnificada !== undefined ?
    'Col ' + cols.accionUnificada + ' ("' + encabezados[cols.accionUnificada] + '")' +
    (cols.accionAutoDetectada ? ' [AUTO-DETECTADA por escaneo]' : '') :
    '❌ NO ENCONTRADA') + '\n';
  deteccion += 'ingreso (viejo): ' + (cols.ingreso !== undefined ? 'Col ' + cols.ingreso : 'NO') + '\n';
  deteccion += 'egreso (viejo): ' + (cols.egreso !== undefined ? 'Col ' + cols.egreso : 'NO') + '\n';
  deteccion += 'uuid: ' + (cols.uuid !== undefined ? 'Col ' + cols.uuid : 'NO') + '\n';

  // Mostrar primeras 5 filas de ejemplo
  var ejemplos = '\n--- PRIMERAS 5 FILAS ---\n';
  for (var f = 1; f < Math.min(6, datos.length); f++) {
    var participante = obtenerParticipanteFila(datos[f], cols) || '?';
    var tipoReg = obtenerTipoRegistro(datos[f], cols);
    var fechaStr = cols.start !== undefined ? datos[f][cols.start] : '?';
    var accionRaw = cols.accionUnificada !== undefined ? datos[f][cols.accionUnificada] : '(formato viejo)';
    var tipoTxt = tipoReg.esIngreso ? '🟢 Entrada' : (tipoReg.esEgreso ? '🔴 Salida' : '⚪ ???');
    if (tipoReg.esTerapia) tipoTxt += '+Terapia';
    if (tipoReg.esComputacion) tipoTxt += '+Compu';
    if (tipoReg.esPermiso) tipoTxt += '+Permiso';
    ejemplos += f + ': ' + participante + ' | "' + accionRaw + '" → ' + tipoTxt + '\n';
  }

  // Contar tipos
  var conteo = { entrada: 0, salida: 0, terapia: 0, computacion: 0, permiso: 0, sinTipo: 0, total: datos.length - 1 };
  for (var f = 1; f < datos.length; f++) {
    var t = obtenerTipoRegistro(datos[f], cols);
    if (t.esIngreso) conteo.entrada++;
    else if (t.esEgreso) conteo.salida++;
    else conteo.sinTipo++;
    if (t.esTerapia) conteo.terapia++;
    if (t.esComputacion) conteo.computacion++;
    if (t.esPermiso) conteo.permiso++;
  }
  var resumen = '\n--- RESUMEN ---\n';
  resumen += 'Total filas: ' + conteo.total + '\n';
  resumen += '🟢 Entradas: ' + conteo.entrada + '\n';
  resumen += '🔴 Salidas: ' + conteo.salida + '\n';
  resumen += '🧘 Terapias: ' + conteo.terapia + '\n';
  resumen += '💻 Computación: ' + conteo.computacion + '\n';
  resumen += '📝 Permisos: ' + conteo.permiso + '\n';
  if (conteo.sinTipo > 0) resumen += '⚠️ Sin tipo: ' + conteo.sinTipo + '\n';

  if (conteo.entrada === 0 && conteo.salida === 0) {
    resumen += '\n❌ PROBLEMA: No se detectaron entradas ni salidas.\n';
    resumen += 'Revisa que la columna de acción tenga valores como:\n';
    resumen += '"Entrada", "Salida" (también reconoce: Ingreso, Egreso, 🟢, 🔴)\n';
  } else {
    resumen += '\n✅ Datos listos para generar reportes.\n';
  }

  SpreadsheetApp.getUi().alert(listaEnc + deteccion + ejemplos + resumen);
}

// ==================== FUNCIONES AUXILIARES ====================
function esMismaFecha(fecha1, fecha2) {
  return fecha1.getFullYear() === fecha2.getFullYear() &&
         fecha1.getMonth() === fecha2.getMonth() &&
         fecha1.getDate() === fecha2.getDate();
}

function estaEnSemana(fecha, inicioSemana) {
  var finSemana = new Date(inicioSemana);
  finSemana.setDate(finSemana.getDate() + 6);
  return fecha >= inicioSemana && fecha <= finSemana;
}

function estaEnMes(fecha, mesFecha) {
  return fecha.getFullYear() === mesFecha.getFullYear() &&
         fecha.getMonth() === mesFecha.getMonth();
}

function validarEnRango(tipo, fechaRegistro, fechaInicio, fechaFin) {
  if (tipo === 'dia' && fechaInicio) {
    return esMismaFecha(fechaRegistro, fechaInicio);
  } else if (tipo === 'semana' && fechaInicio) {
    return estaEnSemana(fechaRegistro, fechaInicio);
  } else if (tipo === 'mes' && fechaInicio) {
    return estaEnMes(fechaRegistro, fechaInicio);
  } else if (tipo === 'rango' && fechaInicio && fechaFin) {
    var fechaReg = new Date(fechaRegistro);
    fechaReg.setHours(0, 0, 0, 0);
    return fechaReg >= fechaInicio && fechaReg <= fechaFin;
  } else if (tipo === 'todo') {
    return true;
  }
  return true;
}

function obtenerTextoPeriodo(tipo, fechaInicio, fechaFin) {
  if (tipo === 'dia') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (tipo === 'semana') {
    var finSemana = new Date(fechaInicio);
    finSemana.setDate(finSemana.getDate() + 6);
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
           ' al ' + Utilities.formatDate(finSemana, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else if (tipo === 'mes') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'MMMM yyyy').toUpperCase();
  } else if (tipo === 'rango') {
    return Utilities.formatDate(fechaInicio, Session.getScriptTimeZone(), 'dd/MM/yyyy') +
           ' al ' + Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } else {
    return 'TODOS LOS REGISTROS';
  }
}
