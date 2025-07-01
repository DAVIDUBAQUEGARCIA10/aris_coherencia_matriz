function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle("ARIS Coherencia de la matriz")
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .addMetaTag("viewport", "width=device-width, initial-scale=1")
  .setFaviconUrl("https://raw.githubusercontent.com/DAVEUBAQUE1996/dave/main/icon-bolivar-conmigo.png")
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTreeData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Comerciales");
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const tree = {};
  const contadorFilas = {};      // # de filas por compañía
  const sumatoriaCoherencia = {}; // sumatoria total por compañía

  data.forEach(row => {
    const [
      compania, riesgo, causa, control,
      calificacionCausa, sujeto, accion,
      objeto, contexto, factores, justificacion,
      calificacionRiesgo, calificacionControl
    ] = row;

    if (!compania || !riesgo || !causa || !control) return;

    const califCausaNum = Number(calificacionCausa) || 0;
    const califRiesgoNum = Number(calificacionRiesgo) || 0;
    const califControlNum = Number(calificacionControl) || 0;
    const sumaFila = califCausaNum + califRiesgoNum + califControlNum;

    // Inicializa estructura si no existe
    if (!tree[compania]) {
      tree[compania] = {
        coherencia: 0,
        riesgos: {}
      };
      contadorFilas[compania] = 0;
      sumatoriaCoherencia[compania] = 0;
    }

    contadorFilas[compania]++;
    sumatoriaCoherencia[compania] += sumaFila;

    if (!tree[compania].riesgos[riesgo]) {
      tree[compania].riesgos[riesgo] = {
        calificacionRiesgo: calificacionRiesgo,
        causas: {}
      };
    }

    if (!tree[compania].riesgos[riesgo].causas[causa]) {
      tree[compania].riesgos[riesgo].causas[causa] = {
        detalles: {
          "Sujeto": sujeto,
          "Acción": accion,
          "Objeto determinado": objeto,
          "Contexto delictivo específico": contexto,
          "4 factores de riesgo": factores,
          "Justificación": justificacion,
          "Calificación Causa": calificacionCausa
        },
        controles: []
      };
    }

    tree[compania].riesgos[riesgo].causas[causa].controles.push({
      texto: control,
      calificacion: calificacionControl
    });
  });

  // Calcular coherencia por compañía (%)
  for (const comp in tree) {
    const total = contadorFilas[comp] * 9;
    const porcentaje = total ? (sumatoriaCoherencia[comp] / total) * 100 : 0;
    tree[comp].coherencia = Math.round(porcentaje);
  }

  return tree;
}
