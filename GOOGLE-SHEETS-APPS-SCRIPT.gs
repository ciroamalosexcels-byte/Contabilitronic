/**
 * PUENTE MI CONTABILIDAD <-> GOOGLE SHEETS
 *
 * 1) Abrí la planilla > Extensiones > Apps Script.
 * 2) Reemplazá todo el código por este archivo.
 * 3) Implementar > Nueva implementación > Aplicación web.
 * 4) Ejecutar como: Yo. Acceso: Cualquier persona.
 * 5) Copiá la URL que termina en /exec y pegala en la app.
 */

const SHEET_DATA = 'DatosApp';

function doGet() {
  try {
    const state = readState_();
    return json_({
      ok: true,
      movs: Array.isArray(state.movs) ? state.movs : [],
      fijos: Array.isArray(state.fijos) ? state.fijos : [],
      cats: Array.isArray(state.cats) ? state.cats : [],
      goalMin: Number(state.goalMin) || 0,
      goalMax: Number(state.goalMax) || 0,
      useFiscal: Boolean(state.useFiscal)
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No se recibieron datos.');
    }
    const state = JSON.parse(e.postData.contents);
    if (!state || !Array.isArray(state.movs) || !Array.isArray(state.fijos)) {
      throw new Error('El formato recibido no es válido.');
    }
    writeState_(state);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function readState_() {
  const sh = getDataSheet_();
  const raw = sh.getRange('A2').getDisplayValue();
  if (!raw) return { movs: [], fijos: [], cats: [], goalMin: 0, goalMax: 0, useFiscal: false };
  return JSON.parse(raw);
}

function writeState_(state) {
  const sh = getDataSheet_();
  sh.getRange('A1').setValue('JSON de Mi Contabilidad — no editar manualmente');
  sh.getRange('A2').setValue(JSON.stringify(state));
  sh.getRange('B1').setValue('Última sincronización');
  sh.getRange('B2').setValue(new Date());
}

function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_DATA);
  if (!sh) sh = ss.insertSheet(SHEET_DATA);
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
