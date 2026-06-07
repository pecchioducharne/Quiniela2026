// ============================================================
// Quiniela 2026 — Sync Supabase → Google Sheets
// ============================================================
// INSTRUCCIONES:
//   1. Crea un Google Sheet nuevo (sheets.google.com)
//   2. Extensiones → Apps Script → pega TODO este código
//   3. Guarda (Ctrl+S) y luego clic en ▶ "Ejecutar" → sincronizar
//   4. Autoriza los permisos que solicite Google
//   5. Para sync automático: Disparadores (reloj) → Añadir disparador
//      → sincronizar → Basado en tiempo → Cada hora o cada día
// ============================================================

const SUPA_URL = 'https://dwqvchnvrdkyeksdebho.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImR3cXZjaG52cmRreWVrc2RlYmhvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzkzNjY4NTMsImV4cCI6MjA5NDk0Mjg1M30.8HfPAhVeZynKj9uj9c7RyYrSsdoHxHkudG3jYz6A9sE';

const GRUPOS = ['A','B','C','D','E','F','G','H','I','J','K','L'];

// Partidos R32 (16avos de final) en orden de pantalla
const R32 = [
  {id:'p74',  lbl:'P74 1°E vs 3er'},  {id:'p77',  lbl:'P77 1°I vs 3er'},
  {id:'p73',  lbl:'P73 2°A vs 2°B'}, {id:'p75',  lbl:'P75 1°F vs 2°C'},
  {id:'p83',  lbl:'P83 2°K vs 2°L'}, {id:'p84',  lbl:'P84 1°H vs 2°J'},
  {id:'p81',  lbl:'P81 1°D vs 3er'}, {id:'p82',  lbl:'P82 1°G vs 3er'},
  {id:'p78',  lbl:'P78 2°E vs 2°I'}, {id:'p76',  lbl:'P76 1°C vs 2°F'},
  {id:'p79',  lbl:'P79 1°A vs 3er'}, {id:'p80',  lbl:'P80 1°L vs 3er'},
  {id:'p88',  lbl:'P88 2°D vs 2°G'}, {id:'p86',  lbl:'P86 1°J vs 2°H'},
  {id:'p85',  lbl:'P85 1°B vs 3er'}, {id:'p87',  lbl:'P87 1°K vs 3er'}
];
const QF = [
  {id:'qf_p89', lbl:'8vos P89'}, {id:'qf_p90', lbl:'8vos P90'},
  {id:'qf_p93', lbl:'8vos P93'}, {id:'qf_p94', lbl:'8vos P94'},
  {id:'qf_p91', lbl:'8vos P91'}, {id:'qf_p92', lbl:'8vos P92'},
  {id:'qf_p95', lbl:'8vos P95'}, {id:'qf_p96', lbl:'8vos P96'}
];
const CF = [
  {id:'cf_p97',  lbl:'4tos P97'},  {id:'cf_p98',  lbl:'4tos P98'},
  {id:'cf_p99',  lbl:'4tos P99'},  {id:'cf_p100', lbl:'4tos P100'}
];
const SF = [
  {id:'sf_p101', lbl:'Semi P101'}, {id:'sf_p102', lbl:'Semi P102'}
];
const FINALES = [
  {id:'p103', lbl:'3er Lugar'}, {id:'p104', lbl:'FINAL / Campeón'}
];

// ── FUNCIÓN PRINCIPAL ──────────────────────────────────────────
function sincronizar() {
  const resp = UrlFetchApp.fetch(
    `${SUPA_URL}/rest/v1/quinelas?select=*&order=created_at.asc`,
    {
      method: 'GET',
      headers: {
        'apikey': SUPA_KEY,
        'Authorization': 'Bearer ' + SUPA_KEY,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    }
  );

  if (resp.getResponseCode() !== 200) {
    throw new Error('Error Supabase ' + resp.getResponseCode() + ': ' + resp.getContentText());
  }

  const filas = JSON.parse(resp.getContentText());
  const participantes = filas.filter(r => r.name !== '__real__');
  const real = filas.find(r => r.name === '__real__') || null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ahora = Utilities.formatDate(new Date(), 'America/Bogota', 'dd/MM/yyyy HH:mm');

  _escribirRespaldo(ss, filas, ahora);
  _escribirGrupos(ss, participantes, real, ahora);
  _escribirEliminatoria(ss, participantes, real, ahora);

  SpreadsheetApp.getActive().toast(
    `✅ ${participantes.length} quinielas sincronizadas · ${ahora}`,
    'Quiniela 2026', 6
  );
}

// ── Hoja 1: RESPALDO (JSON crudo — backup completo) ────────────
function _escribirRespaldo(ss, filas, ahora) {
  let sh = ss.getSheetByName('📋 Respaldo');
  if (!sh) sh = ss.insertSheet('📋 Respaldo', 0);
  sh.clearContents();
  sh.clearFormats();

  const hdrs = ['Nombre', 'Datos (JSON completo)', 'Fecha creación', 'ID'];
  const hRow = sh.getRange(1, 1, 1, hdrs.length);
  hRow.setValues([hdrs])
      .setBackground('#1a3a5c')
      .setFontColor('#ffffff')
      .setFontWeight('bold');

  if (filas.length > 0) {
    const datos = filas.map(r => [
      r.name,
      JSON.stringify(r.data || {}),
      r.created_at ? Utilities.formatDate(new Date(r.created_at), 'America/Bogota', 'dd/MM/yyyy HH:mm') : '',
      r.id || ''
    ]);
    sh.getRange(2, 1, datos.length, hdrs.length).setValues(datos);

    // Resaltar fila __real__
    const realFila = filas.findIndex(r => r.name === '__real__');
    if (realFila >= 0) {
      sh.getRange(realFila + 2, 1, 1, hdrs.length)
        .setBackground('#fef3c7').setFontWeight('bold');
    }
  }

  sh.getRange(sh.getLastRow() + 2, 1)
    .setValue('Última sincronización: ' + ahora)
    .setFontStyle('italic').setFontColor('#718096');

  sh.autoResizeColumn(1);
  sh.setColumnWidth(2, 500);
  sh.setFrozenRows(1);
}

// ── Hoja 2: GRUPOS (tabla legible) ─────────────────────────────
function _escribirGrupos(ss, participantes, real, ahora) {
  let sh = ss.getSheetByName('⚽ Grupos');
  if (!sh) sh = ss.insertSheet('⚽ Grupos', 1);
  sh.clearContents();
  sh.clearFormats();

  // Cabecera: Nombre | A1° A2° A3° A4° | B1° B2° B3° B4° | ... | Terceros
  const colsGrupo = GRUPOS.flatMap(g => [`${g} 1°`, `${g} 2°`, `${g} 3°`, `${g} 4°`]);
  const hdrs = ['Nombre', ...colsGrupo, 'Terceros clasificados'];
  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');

  // Colorear columnas por grupo (alternado verde / azul oscuro)
  const colores = ['#276749', '#1a3a5c'];
  GRUPOS.forEach((g, gi) => {
    sh.getRange(1, 2 + gi * 4, 1, 4).setBackground(colores[gi % 2]);
  });

  // Filas de participantes
  const filas = participantes.map(p => _filaGrupos(p));
  if (filas.length > 0) {
    sh.getRange(2, 1, filas.length, hdrs.length).setValues(filas);
    // Franjas zebra
    filas.forEach((_, i) => {
      if (i % 2 === 0) sh.getRange(i + 2, 1, 1, hdrs.length).setBackground('#f8fafc');
    });
  }

  // Fila resultado real
  if (real) {
    const realFila = _filaGrupos(real);
    realFila[0] = '🏆 RESULTADO REAL';
    const row = 2 + filas.length;
    sh.getRange(row, 1, 1, hdrs.length)
      .setValues([realFila])
      .setBackground('#d69e2e').setFontColor('#fff').setFontWeight('bold');
  }

  sh.getRange(sh.getLastRow() + 2, 1)
    .setValue('Última sincronización: ' + ahora)
    .setFontStyle('italic').setFontColor('#718096');

  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.autoResizeColumn(1);
}

function _filaGrupos(p) {
  const d = p.data || {};
  const grupos = d.groups || {};
  const terceros = (d.thirds || []).sort().join(', ');
  const cols = GRUPOS.flatMap(g => {
    const orden = grupos[g] || [null, null, null, null];
    return [orden[0] || '', orden[1] || '', orden[2] || '', orden[3] || ''];
  });
  return [p.name, ...cols, terceros];
}

// ── Hoja 3: ELIMINATORIA (bracket picks) ───────────────────────
function _escribirEliminatoria(ss, participantes, real, ahora) {
  let sh = ss.getSheetByName('🏆 Eliminatoria');
  if (!sh) sh = ss.insertSheet('🏆 Eliminatoria', 2);
  sh.clearContents();
  sh.clearFormats();

  const partidos = [...R32, ...QF, ...CF, ...SF, ...FINALES, {id:'scorer', lbl:'Goleador'}];
  const hdrs = ['Nombre', ...partidos.map(m => m.lbl)];
  sh.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
    .setBackground('#1a3a5c').setFontColor('#fff').setFontWeight('bold');

  // Colorear secciones de rondas
  const secciones = [
    {start:2,  len:16, bg:'#1a3a5c'}, // R32
    {start:18, len:8,  bg:'#276749'}, // 8vos
    {start:26, len:4,  bg:'#744210'}, // 4tos
    {start:30, len:2,  bg:'#553c9a'}, // Semi
    {start:32, len:2,  bg:'#c53030'}, // Finales
    {start:34, len:1,  bg:'#d69e2e'}  // Goleador
  ];
  secciones.forEach(s => {
    if (s.start <= hdrs.length) {
      sh.getRange(1, s.start, 1, Math.min(s.len, hdrs.length - s.start + 1))
        .setBackground(s.bg);
    }
  });

  const filas = participantes.map(p => _filaEliminatoria(p, partidos));
  if (filas.length > 0) {
    sh.getRange(2, 1, filas.length, hdrs.length).setValues(filas);
    filas.forEach((_, i) => {
      if (i % 2 === 0) sh.getRange(i + 2, 1, 1, hdrs.length).setBackground('#f8fafc');
    });
  }

  if (real) {
    const realFila = _filaEliminatoria(real, partidos);
    realFila[0] = '🏆 RESULTADO REAL';
    const row = 2 + filas.length;
    sh.getRange(row, 1, 1, hdrs.length)
      .setValues([realFila])
      .setBackground('#d69e2e').setFontColor('#fff').setFontWeight('bold');
  }

  sh.getRange(sh.getLastRow() + 2, 1)
    .setValue('Última sincronización: ' + ahora)
    .setFontStyle('italic').setFontColor('#718096');

  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);
  sh.autoResizeColumn(1);
}

function _filaEliminatoria(p, partidos) {
  const d = p.data || {};
  const picks = d.picks || {};
  const goleador = d.goleador || '';
  const cols = partidos.map(m => {
    if (m.id === 'scorer') return goleador;
    const w = picks[m.id];
    return w ? `${w.f || ''} ${w.n || ''}`.trim() : '';
  });
  return [p.name, ...cols];
}

// ── TRIGGER AUTOMÁTICO ─────────────────────────────────────────
// Ejecuta esta función UNA SOLA VEZ para instalar el trigger diario
function instalarTriggerDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizar')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sincronizar')
    .timeBased()
    .everyDays(1)
    .atHour(11) // 6am hora Colombia (UTC-5)
    .create();

  SpreadsheetApp.getActive().toast(
    '⏰ Sync automático instalado: todos los días a las 6am (Colombia)',
    'Trigger OK', 5
  );
}

// Para sync cada hora (opcional — comenta el trigger diario si usas este):
function instalarTriggerHorario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizar')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sincronizar')
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getActive().toast(
    '⏰ Sync automático instalado: cada hora',
    'Trigger OK', 5
  );
}
