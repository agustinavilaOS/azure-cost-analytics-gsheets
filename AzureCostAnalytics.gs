/**
 * Función especial que se ejecuta automáticamente cuando se abre la hoja de cálculo.
 * Crea un menú personalizado en la interfaz de Google Sheets para ejecutar los reportes.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Reportes de Impacto')
    .addItem('Paso 1: ✨ Revelar Top 15 de Crecimiento', 'generarTop15')
    .addItem('Paso 1B: 🟢 Revelar Top 15 de Optimización', 'generarTop15Optimizados')
    .addSeparator()
    .addItem('Paso 2: 📊 Visualizar el Éxito', 'generarGrafica')
    .addItem('Paso 2B: 🟢 Visualizar Optimización', 'generarGraficaOptimizados')
    .addSeparator()
    .addItem('Paso 3: ✉️ Distribuir Inteligencia', 'enviarGraficaPorCorreo')
    .addToUi();
}

/**
 * Función base para obtener y procesar los datos de crecimiento.
 * Es utilizada por las funciones de generación de reportes para evitar repetir código.
 * @returns {Array} Un arreglo de objetos con los 15 recursos de mayor crecimiento.
 */
function obtenerDatosCrecimiento() {
  const ss = SpreadsheetApp.getActive();
  const hojaMesAComparar = ss.getSheetByName("Mes a comparar");
  const hojaMesBase = ss.getSheetByName("Mes Base");

  if (!hojaMesAComparar || !hojaMesBase) {
    SpreadsheetApp.getUi().alert('Error: No se encontraron las hojas "Mes a comparar" o "Mes Base". Por favor, verifica que existan.');
    return null;
  }

  const datos = hojaMesAComparar.getDataRange().getValues();
  const MesBaseDatos = hojaMesBase.getDataRange().getValues();

  const headers = datos[0];
  const idxResourceId = headers.indexOf("ResourceId");
  const idxCostUSD = headers.indexOf("CostUSD");

  const MesBaseHeaders = MesBaseDatos[0];
  const idxMesBaseId = MesBaseHeaders.indexOf("ResourceId");
  const idxMesBaseCost = MesBaseHeaders.indexOf("CostUSD");

  const actual = {};
  const MesBase = {};

  for (let i = 1; i < datos.length; i++) {
    const id = datos[i][idxResourceId];
    const nombre = id.split("/").pop();
    const costo = parseFloat(datos[i][idxCostUSD]) || 0;
    actual[nombre] = (actual[nombre] || 0) + costo;
  }

  for (let i = 1; i < MesBaseDatos.length; i++) {
    const id = MesBaseDatos[i][idxMesBaseId];
    const nombre = id.split("/").pop();
    const costo = parseFloat(MesBaseDatos[i][idxMesBaseCost]) || 0;
    MesBase[nombre] = (MesBase[nombre] || 0) + costo;
  }

  const comparacion = [];
  for (let nombre in actual) {
    const actualCost = actual[nombre];
    const MesBaseCost = MesBase[nombre] || 0;
    const crecimientoUSD = actualCost - MesBaseCost;
    const crecimientoPct = MesBaseCost === 0 ? 1 : crecimientoUSD / MesBaseCost;

    comparacion.push({
      nombre,
      MesBaseCost,
      actualCost,
      crecimientoUSD,
      crecimientoPct,
    });
  }

  comparacion.sort((a, b) => b.crecimientoUSD - a.crecimientoUSD);
  return comparacion.slice(0, 15);
}

/**
 * Obtiene el Top 15 de recursos que MÁS DISMINUYERON su costo (optimización).
 * @returns {Array} Un arreglo con los 15 recursos de mayor optimización.
 */
function obtenerDatosOptimizacion() {
  const ss = SpreadsheetApp.getActive();
  const hojaMesAComparar = ss.getSheetByName("Mes a comparar");
  const hojaMesBase = ss.getSheetByName("Mes Base");

  if (!hojaMesAComparar || !hojaMesBase) {
    SpreadsheetApp.getUi().alert('Error: No se encontraron las hojas "Mes a comparar" o "Mes Base". Por favor, verifica que existan.');
    return null;
  }

  // Se reutiliza la lógica de comparación general, pero se filtra para solo los que disminuyeron costo.
  const datos = hojaMesAComparar.getDataRange().getValues();
  const MesBaseDatos = hojaMesBase.getDataRange().getValues();

  const headers = datos[0];
  const idxResourceId = headers.indexOf("ResourceId");
  const idxCostUSD = headers.indexOf("CostUSD");

  const MesBaseHeaders = MesBaseDatos[0];
  const idxMesBaseId = MesBaseHeaders.indexOf("ResourceId");
  const idxMesBaseCost = MesBaseHeaders.indexOf("CostUSD");

  const actual = {};
  const MesBase = {};

  for (let i = 1; i < datos.length; i++) {
    const id = datos[i][idxResourceId];
    const nombre = id.split("/").pop();
    const costo = parseFloat(datos[i][idxCostUSD]) || 0;
    actual[nombre] = (actual[nombre] || 0) + costo;
  }

  for (let i = 1; i < MesBaseDatos.length; i++) {
    const id = MesBaseDatos[i][idxMesBaseId];
    const nombre = id.split("/").pop();
    const costo = parseFloat(MesBaseDatos[i][idxMesBaseCost]) || 0;
    MesBase[nombre] = (MesBase[nombre] || 0) + costo;
  }

  const comparacion = [];
  for (let nombre in actual) {
    const actualCost = actual[nombre];
    const MesBaseCost = MesBase[nombre] || 0;
    const crecimientoUSD = actualCost - MesBaseCost;
    const crecimientoPct = MesBaseCost === 0 ? 1 : crecimientoUSD / MesBaseCost;

    comparacion.push({
      nombre,
      MesBaseCost,
      actualCost,
      crecimientoUSD,
      crecimientoPct,
    });
  }

  // Solo optimizados (crecimientoUSD negativo)
  const soloOptimizados = comparacion.filter(r => r.crecimientoUSD < 0);
  soloOptimizados.sort((a, b) => a.crecimientoUSD - b.crecimientoUSD); // Más negativo primero (más ahorro)
  return soloOptimizados.slice(0, 15);
}

/**
 * PASO 1: Genera la tabla con el Top 15 de recursos con mayor crecimiento.
 */
function generarTop15() {
  const top15 = obtenerDatosCrecimiento();
  if (!top15) return; // Detener si hubo un error en la obtención de datos.

  const hojaCrecimiento = getOrCreateSheet("Top15Crecimiento");
  hojaCrecimiento.clear();
  hojaCrecimiento.setHiddenGridlines(true);

  hojaCrecimiento.appendRow(["Recurso", "Base", "Actual", "Crecimiento $", "Crecimiento %"]);

  top15.forEach(r => {
    hojaCrecimiento.appendRow([r.nombre, r.MesBaseCost, r.actualCost, r.crecimientoUSD, r.crecimientoPct]);
  });

  const rows = top15.length + 1;
  const rangeCrecimiento = hojaCrecimiento.getRange(1, 1, rows, 5);
  rangeCrecimiento.setFontFamily("Poppins").setFontSize(12).setFontWeight("normal").setBackground(null).setBorder(true, true, true, true, true, true);
  hojaCrecimiento.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9d9d9");

  const totalFila = rows + 1;
  hojaCrecimiento.getRange(totalFila, 1, 1, 5).setFontFamily("Poppins").setFontSize(12).setFontWeight("bold").setBackground("#d9d9d9");
  hojaCrecimiento.getRange(totalFila, 1).setValue("TOTAL");
  hojaCrecimiento.getRange(totalFila, 2).setFormula(`=SUM(B2:B${rows})`);
  hojaCrecimiento.getRange(totalFila, 3).setFormula(`=SUM(C2:C${rows})`);
  hojaCrecimiento.getRange(totalFila, 4).setFormula(`=SUM(D2:D${rows})`);
  hojaCrecimiento.getRange(totalFila, 5).setValue("");

  hojaCrecimiento.getRange(`A2:E${totalFila}`).setWrap(false);
  hojaCrecimiento.getRange(`B2:D${totalFila}`).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  hojaCrecimiento.getRange(`E2:E${totalFila}`).setNumberFormat("0.00%").setHorizontalAlignment("right");

  hojaCrecimiento.setColumnWidth(1, 535);
  for (let col = 2; col <= 5; col++) {
    hojaCrecimiento.setColumnWidth(col, 135);
  }
  SpreadsheetApp.getUi().alert('¡Éxito!', 'La tabla con el Top 15 de crecimiento ha sido generada.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * PASO 1B: Genera la tabla con el Top 15 de recursos que más OPTIMIZARON (ahorro).
 */
function generarTop15Optimizados() {
  const top15Optimizados = obtenerDatosOptimizacion();
  if (!top15Optimizados || top15Optimizados.length === 0) {
    SpreadsheetApp.getUi().alert('No hay recursos con disminución de costo significativos este mes.');
    return;
  }

  const hojaOptimizados = getOrCreateSheet("Top15Optimizacion");
  hojaOptimizados.clear();
  hojaOptimizados.setHiddenGridlines(true);

  hojaOptimizados.appendRow(["Recurso", "Base", "Actual", "Ahorro $", "Reducción %"]);

  top15Optimizados.forEach(r => {
    hojaOptimizados.appendRow([r.nombre, r.MesBaseCost, r.actualCost, r.crecimientoUSD, r.crecimientoPct]);
  });

  const rows = top15Optimizados.length + 1;
  const range = hojaOptimizados.getRange(1, 1, rows, 5);
  range.setFontFamily("Poppins").setFontSize(12).setFontWeight("normal").setBackground(null).setBorder(true, true, true, true, true, true);
  hojaOptimizados.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#d9d9d9");

  const totalFila = rows + 1;
  hojaOptimizados.getRange(totalFila, 1, 1, 5).setFontFamily("Poppins").setFontSize(12).setFontWeight("bold").setBackground("#d9d9d9");
  hojaOptimizados.getRange(totalFila, 1).setValue("TOTAL");
  hojaOptimizados.getRange(totalFila, 2).setFormula(`=SUM(B2:B${rows})`);
  hojaOptimizados.getRange(totalFila, 3).setFormula(`=SUM(C2:C${rows})`);
  hojaOptimizados.getRange(totalFila, 4).setFormula(`=SUM(D2:D${rows})`);
  hojaOptimizados.getRange(totalFila, 5).setValue("");

  hojaOptimizados.getRange(`A2:E${totalFila}`).setWrap(false);
  hojaOptimizados.getRange(`B2:D${totalFila}`).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  hojaOptimizados.getRange(`E2:E${totalFila}`).setNumberFormat("0.00%").setHorizontalAlignment("right");

  hojaOptimizados.setColumnWidth(1, 535);
  for (let col = 2; col <= 5; col++) {
    hojaOptimizados.setColumnWidth(col, 135);
  }
  SpreadsheetApp.getUi().alert('¡Éxito!', 'La tabla con el Top 15 de optimización ha sido generada.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * PASO 2: Genera la tabla y el gráfico de pastel en una nueva hoja (Crecimiento).
 */
function generarGrafica() {
  const top15 = obtenerDatosCrecimiento();
  if (!top15) return;

  const hojaGrafica = getOrCreateSheet("GraficaCrecimiento");
  hojaGrafica.clear();
  hojaGrafica.setHiddenGridlines(true);

  // Elimina gráficos anteriores para evitar duplicados
  hojaGrafica.getCharts().forEach(chart => hojaGrafica.removeChart(chart));

  hojaGrafica.appendRow(["Recurso", "Crecimiento $"]);
  top15.forEach(r => {
    hojaGrafica.appendRow([r.nombre, r.crecimientoUSD]);
  });

  const filasGraf = top15.length + 1;
  const rangoGraf = hojaGrafica.getRange(1, 1, filasGraf, 2);
  rangoGraf.setFontFamily("Poppins").setFontSize(12).setFontWeight("normal").setBackground(null).setBorder(true, true, true, true, true, true);

  hojaGrafica.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#d9d9d9");

  const totalFilaGraf = filasGraf + 1;
  hojaGrafica.getRange(totalFilaGraf, 1, 1, 2).setFontFamily("Poppins").setFontSize(12).setFontWeight("bold").setBackground("#d9d9d9");
  hojaGrafica.getRange(totalFilaGraf, 1).setValue("TOTAL");
  hojaGrafica.getRange(totalFilaGraf, 2).setFormula(`=SUM(B2:B${filasGraf})`);

  hojaGrafica.getRange(`A2:B${totalFilaGraf}`).setWrap(false);
  hojaGrafica.setColumnWidth(1, 535);
  hojaGrafica.setColumnWidth(2, 135);
  hojaGrafica.getRange(`B2:B${totalFilaGraf}`).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");

  const rangoDatosGrafico = hojaGrafica.getRange(1, 1, filasGraf, 2);
  const chartBuilder = hojaGrafica.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(rangoDatosGrafico)
    .setOption('pieHole', 0.75)
    .setPosition(1, 3, 0, 0)
    .setOption('width', 700)
    .setOption('height', 360)
    .setOption('backgroundColor', 'none')
    .setOption('is3D', true)
    .setOption('title', 'Recursos en Crecimiento')
    .setOption('titleTextStyle', { color: 'black', bold: true, fontName: 'Sans Serif' })
    .setOption('pieSliceText', 'value-and-percentage')
    .setOption('pieSliceTextStyle', { fontSize: 12, fontName: 'Sans Serif' })
    .setOption('legend', { position: 'labeled', textStyle: { fontSize: 12, fontName: 'Sans Serif' } })
    .build();

  hojaGrafica.insertChart(chartBuilder);
  SpreadsheetApp.getUi().alert('¡Visualización Creada!', 'El gráfico de impacto ha sido generado exitosamente.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * PASO 2B: Genera la tabla y el gráfico de pastel en una nueva hoja (Optimización).
 */
function generarGraficaOptimizados() {
  const top15Optimizados = obtenerDatosOptimizacion();
  if (!top15Optimizados || top15Optimizados.length === 0) {
    SpreadsheetApp.getUi().alert('No hay recursos con disminución de costo significativos este mes.');
    return;
  }

  const hojaGraficaOpt = getOrCreateSheet("GraficaOptimizacion");
  hojaGraficaOpt.clear();
  hojaGraficaOpt.setHiddenGridlines(true);

  // Elimina gráficos anteriores
  hojaGraficaOpt.getCharts().forEach(chart => hojaGraficaOpt.removeChart(chart));

  hojaGraficaOpt.appendRow(["Recurso", "Ahorro $"]);
  top15Optimizados.forEach(r => {
    hojaGraficaOpt.appendRow([r.nombre, Math.abs(r.crecimientoUSD)]); // Mostrar el ahorro como valor positivo
  });

  const filas = top15Optimizados.length + 1;
  const rango = hojaGraficaOpt.getRange(1, 1, filas, 2);
  rango.setFontFamily("Poppins").setFontSize(12).setFontWeight("normal").setBackground(null).setBorder(true, true, true, true, true, true);

  hojaGraficaOpt.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#d9d9d9");

  const totalFila = filas + 1;
  hojaGraficaOpt.getRange(totalFila, 1, 1, 2).setFontFamily("Poppins").setFontSize(12).setFontWeight("bold").setBackground("#d9d9d9");
  hojaGraficaOpt.getRange(totalFila, 1).setValue("TOTAL");
  hojaGraficaOpt.getRange(totalFila, 2).setFormula(`=SUM(B2:B${filas})`);

  hojaGraficaOpt.getRange(`A2:B${totalFila}`).setWrap(false);
  hojaGraficaOpt.setColumnWidth(1, 535);
  hojaGraficaOpt.setColumnWidth(2, 135);
  hojaGraficaOpt.getRange(`B2:B${totalFila}`).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");

  const rangoDatosGrafico = hojaGraficaOpt.getRange(1, 1, filas, 2);
  const chartBuilder = hojaGraficaOpt.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(rangoDatosGrafico)
    .setOption('pieHole', 0.75)
    .setPosition(1, 3, 0, 0)
    .setOption('width', 700)
    .setOption('height', 360)
    .setOption('backgroundColor', 'none')
    .setOption('is3D', true)
    .setOption('title', 'Recursos con Mayor Optimización (Ahorro)')
    .setOption('titleTextStyle', { color: 'black', bold: true, fontName: 'Sans Serif' })
    .setOption('pieSliceText', 'value-and-percentage')
    .setOption('pieSliceTextStyle', { fontSize: 12, fontName: 'Sans Serif' })
    .setOption('legend', { position: 'labeled', textStyle: { fontSize: 12, fontName: 'Sans Serif' } })
    .build();

  hojaGraficaOpt.insertChart(chartBuilder);
  SpreadsheetApp.getUi().alert('¡Visualización de optimización creada!', 'El gráfico de optimización ha sido generado exitosamente.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * PASO 3: Envía por correo electrónico un PDF de la hoja del gráfico (crecimiento).
 * Si quieres el envío del gráfico de optimización, deberías duplicar esta función cambiando la hoja objetivo.
 */
function enviarGraficaPorCorreo() {
  const ss = SpreadsheetApp.getActive();
  const hojaGrafica = ss.getSheetByName("GraficaCrecimiento");

  if (!hojaGrafica) {
    SpreadsheetApp.getUi().alert('¡Acción Requerida!', 'Primero debes generar el gráfico usando el "Paso 2: Visualizar el Éxito".', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const ui = SpreadsheetApp.getUi();
  const emailResponse = ui.prompt(
    'Confirmar Envío de Reporte',
    'Por favor, ingresa el correo del destinatario:',
    ui.ButtonSet.OK_CANCEL);

  if (emailResponse.getSelectedButton() != ui.Button.OK) {
    return; // El usuario canceló
  }

  const destinatario = emailResponse.getResponseText();
  if (!destinatario) {
    ui.alert('Error', 'No se ingresó un correo electrónico.', ui.ButtonSet.OK);
    return;
  }

  const url = ss.getUrl().replace(/edit$/, '');
  const gid = hojaGrafica.getSheetId();
  const exportUrl = `${url}export?exportFormat=pdf&format=pdf&gid=${gid}&size=A4&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenum=UNDEFINED&gridlines=false&fzr=false`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  const pdfBlob = response.getBlob().setName("GraficaCrecimiento.pdf");

  const asunto = "Reporte de Impacto - Crecimiento de Recursos Azure";
  const cuerpo = `
Hola,

Adjunto encontrarás el reporte visual de los 15 recursos con mayor crecimiento en costos de Azure. Este análisis te permitirá tomar decisiones estratégicas informadas.

¡Que tengas un día productivo!

Saludos cordiales,
Tu Asistente de Reportes Automatizados
  `;

  GmailApp.sendEmail(destinatario, asunto, cuerpo, {
    attachments: [pdfBlob],
    name: "Reporteador Automático Azure"
  });

  ui.alert('¡Misión Cumplida!', `El reporte ha sido enviado exitosamente a ${destinatario}.`, ui.ButtonSet.OK);
}

/**
 * Función de utilidad para obtener una hoja por su nombre.
 * Si la hoja no existe, la crea.
 * @param {string} name El nombre de la hoja a obtener o crear.
 * @returns {Sheet} El objeto de la hoja de cálculo.
 */
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}
