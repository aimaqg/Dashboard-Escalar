// scripts/sync-finanzas.js
// Descarga los archivos EEFF-*.xls(x) de la carpeta de SharePoint (vía Microsoft Graph),
// extrae KPIs financieros por unidad y los guarda en finanzas.json
//
// Soporta dos familias de formato:
//   A) Ejecución presupuestal mensualizada (Charlotte, Complex, Martinique):
//      hoja con columnas ENERO..DICIEMBRE + PTTO acumulado + EJECUCIÓN
//   B) Ejecución anual + Estado de Resultados (Oz):
//      hoja EJECUCION con presupuesto anual, hoja ESTADO DE RESULTADOS con col ACUMULADO
//
// Uso normal (GitHub Actions):  node scripts/sync-finanzas.js
// Uso local (probar parser):    node scripts/sync-finanzas.js --local archivo1.xls archivo2.xlsx ...

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const TENANT = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// === CONFIGURACIÓN DE SHAREPOINT ===
const SHAREPOINT_HOST = 'escalarasesorias.sharepoint.com';
const SITE_PATH = '/sites/Escalar';
const EEFF_SHARE_URL = 'https://escalarasesorias.sharepoint.com/:f:/s/Escalar/IgAFNxZ-bfPtRrpMDL4yJwEIATOW2G4xvwKxEa7phg1yGQc?e=89GeKh';
const FOLDER_PATH_CANDIDATES = ['/Escalar/EEFF', '/EEFF', '/Escalar/Estados Financieros'];
// ===================================

const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

// ── Utilidades ────────────────────────────────────────────────────────────────
function norm(s){ return String(s == null ? '' : s).replace(/\s+/g,' ').trim().toUpperCase(); }
function num(v){
  if(v === null || v === undefined || v === '') return null;
  const n = typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.\-]/g,''));
  return isFinite(n) ? n : null;
}
function tituloCorto(s){
  return String(s||'').replace(/\s+/g,' ').trim()
    .toLowerCase().replace(/(^|\s|\/|-)\S/g, t => t.toUpperCase());
}
function rowsOf(wb, sheetName){
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
}
function findRow(rows, pred, from){
  for(let i = from || 0; i < rows.length; i++){ if(pred(rows[i] || [], i)) return i; }
  return -1;
}
function findRowLast(rows, pred){
  for(let i = rows.length - 1; i >= 0; i--){ if(pred(rows[i] || [], i)) return i; }
  return -1;
}
function cellMatch(row, re){
  for(let j = 0; j < row.length; j++){ if(re.test(norm(row[j]))) return j; }
  return -1;
}
function mesEnTexto(s){
  const n = norm(s);
  for(const m of MESES){ if(n.includes(m)) return m; }
  return null;
}
// Fila cuyo texto (en cualquier celda) empieza por el regex; retorna {i, j}
function findLabelCell(rows, re, from, to){
  for(let i = from || 0; i < (to || rows.length); i++){
    const j = cellMatch(rows[i] || [], re);
    if(j >= 0) return { i, j };
  }
  return null;
}

// ── FORMATO A: ejecución presupuestal mensualizada ───────────────────────────
const RE_SERVICIO = /ACUEDUCTO|ALCANTARILL|ENERG[IÍ]A|ENVIASEO|TASA DE ASEO|TASA ASEO|TEL[EÉ]FONO|CELULAR|INTERNET|CITOFON|^GAS\b|TASA DE SEGURIDAD|IMPUESTO A LA SEGURIDAD/;

function filaMeses(rows){
  return findRow(rows, r => r.filter(c => MESES.some(m => norm(c).startsWith(m))).length >= 6);
}

function parseFormatoA(rows){
  const mIdx = filaMeses(rows);
  if(mIdx < 0) return null;

  const monthCols = {};
  rows[mIdx].forEach((c, j) => {
    const n = norm(c);
    for(const m of MESES){ if(n.startsWith(m) && monthCols[m] === undefined) monthCols[m] = j; }
  });
  const maxMonthCol = Math.max(...Object.values(monthCols));

  // Encabezado "virtual" por columna: concatenación de las celdas del bloque de encabezado
  const virtual = {};
  const nCols = Math.max(...rows.slice(0, mIdx + 2).map(r => (r || []).length), 0);
  for(let j = maxMonthCol + 1; j < nCols; j++){
    const parts = [];
    for(let i = Math.max(0, mIdx - 3); i <= mIdx; i++){
      const c = (rows[i] || [])[j];
      if(c !== null && c !== undefined && String(c).trim()) parts.push(norm(c));
    }
    virtual[j] = parts.join(' ');
  }

  // Columna de presupuesto acumulado a la fecha
  let pptoCol = -1;
  for(const j in virtual){
    const v = virtual[j];
    if(/(PPTO|PTTO|PPTDO)/.test(v) && !/A[ÑN]O|ANUAL/.test(v)){ pptoCol = +j; break; }
  }
  // Columna de ejecución acumulada
  let ejecCol = -1;
  for(const j in virtual){
    const v = virtual[j];
    if(/(EJECUCI|EJECUTADO)/.test(v) && !/DIFEREN|VARIACION|CUMPLIMIENTO|%/.test(v)){ ejecCol = +j; break; }
  }
  if(pptoCol < 0 || ejecCol < 0) return null;

  const label = i => norm((rows[i] || [])[0]);
  const val = (i, j) => num((rows[i] || [])[j]) || 0;

  const iIngresos = findRow(rows, r => /^TOTAL (DE )?(INGRESOS|APORTES)$/.test(norm(r[0])));
  const iGastosTotal = findRowLast(rows, r => /^(GRAN TOTAL (DE )?(GASTOS|EGRESOS)|TOTAL (DE )?(GASTOS|EGRESOS))$/.test(norm(r[0])));
  let iExcedente = findRow(rows, r => /^(EXCEDENTE?S? (O |DEL )|RESULTADO (DEL )?EJERCICIO)/.test(norm(r[0])), mIdx);
  if(iIngresos < 0 || iGastosTotal < 0) throw new Error('Formato A: no se encontraron filas de totales');

  // Mes actual: mes en el encabezado de ppto/ejecución, o último mes con datos en TOTAL INGRESOS
  let mes = mesEnTexto(virtual[pptoCol]) || mesEnTexto(virtual[ejecCol]);
  if(!mes){
    for(const m of MESES){ if(val(iIngresos, monthCols[m]) !== 0) mes = m; }
  }

  const mesesHasta = mes ? MESES.slice(0, MESES.indexOf(mes) + 1) : [];
  const mensual = {
    meses: mesesHasta,
    ingresos: mesesHasta.map(m => val(iIngresos, monthCols[m])),
    gastos: mesesHasta.map(m => val(iGastosTotal, monthCols[m]))
  };

  // Detalle de ingresos: filas antes de TOTAL INGRESOS
  const ingresosDetalle = [];
  function capturarIngreso(i){
    const lb = label(i);
    if(!lb || /^(SUBTOTAL|TOTAL|GRAN TOTAL|CIFRAS?|INGRESOS|APORTES|OTROS INGRESOS)\b/.test(lb)) return;
    if(/^(INGRESOS|APORTES|OTROS INGRESOS)( OPERACIONALES| NO OPERACIONALES)?$/.test(lb)) return;
    const p = num((rows[i]||[])[pptoCol]), e = num((rows[i]||[])[ejecCol]);
    if(p === null && e === null) return;
    const ppto = p || 0, ejec = e || 0;
    if(ppto === 0 && ejec === 0) return;
    ingresosDetalle.push({ rubro: tituloCorto(rows[i][0]), ppto, ejecutado: ejec, diferencia: ejec - ppto });
  }
  for(let i = mIdx + 1; i < iIngresos; i++) capturarIngreso(i);
  // Sección "OTROS INGRESOS" después de gastos (formato Dijon)
  const iOtros = findRow(rows, r => /^OTROS INGRESOS/.test(norm(r[0])), iGastosTotal);
  if(iOtros > 0){
    const iOtrosFin = findRow(rows, r => /^TOTAL OTROS INGRESOS/.test(norm(r[0])), iOtros);
    for(let i = iOtros + 1; i < (iOtrosFin > 0 ? iOtrosFin : rows.length); i++) capturarIngreso(i);
  }

  // Detalle de rubros de gasto (tornado) + servicios públicos mensuales
  const candidatos = [];
  for(let i = iIngresos + 1; i < iGastosTotal; i++){
    const lb = label(i);
    if(!lb) continue;
    if(/^(SUBTOTAL|TOTAL|GRAN TOTAL|CIFRAS?|EXCEDENTE|RESULTADO)/.test(lb)) continue;
    if(lb === 'DIVERSOS') continue; // agrupador con totales propios (formato Charlotte)
    const p = num((rows[i]||[])[pptoCol]), e = num((rows[i]||[])[ejecCol]);
    if(p === null && e === null) continue; // encabezado de sección
    const ppto = p || 0, ejec = e || 0;
    if(ppto === 0 && ejec === 0) continue;
    candidatos.push({ fila: i, lb, rubro: tituloCorto(rows[i][0]), ppto, ejecutado: ejec, diferencia: ejec - ppto });
  }
  // Eliminar filas "agregado" cuyo valor es la suma de las filas hijas contiguas (formato Dijon)
  const esAgregado = new Set();
  const cerca = (a, b, rel) => Math.abs(a - b) <= Math.max(2, Math.abs(a) * rel);
  for(let x = 0; x < candidatos.length; x++){
    let sp = 0, se = 0;
    for(let y = x + 1; y < candidatos.length; y++){
      if(candidatos[y].fila - candidatos[y - 1].fila > 3) break; // bloque no contiguo
      sp += candidatos[y].ppto; se += candidatos[y].ejecutado;
      const pExacto = cerca(candidatos[x].ppto, sp, 0.005), eExacto = cerca(candidatos[x].ejecutado, se, 0.005);
      if((pExacto && cerca(candidatos[x].ejecutado, se, 0.15)) || (eExacto && cerca(candidatos[x].ppto, sp, 0.15))){
        esAgregado.add(x); break;
      }
      if(sp > Math.abs(candidatos[x].ppto) * 1.2 && se > Math.abs(candidatos[x].ejecutado) * 1.2) break;
    }
  }
  const detalle = candidatos.filter((c, x) => !esAgregado.has(x));
  const desviaciones = detalle.map(c => ({ rubro: c.rubro, ppto: c.ppto, ejecutado: c.ejecutado, diferencia: c.diferencia }));
  const servicios = [];
  for(const c of detalle){
    if(RE_SERVICIO.test(c.lb) && !/MTTO|MANTENIMIENTO|SUMINISTRO/.test(c.lb) && mesesHasta.length){
      servicios.push({ nombre: c.rubro, meses: mesesHasta, valores: mesesHasta.map(m => val(c.fila, monthCols[m])) });
    }
  }

  // Composición por categorías: subtotales/totales con nombre dentro de la sección de gastos
  let composicion = [];
  for(let i = iIngresos + 1; i <= iGastosTotal; i++){
    const lb = label(i);
    let nombre = null;
    let m2;
    if((m2 = lb.match(/^(?:SUB)?TOTAL(?:ES)?\s+(.+)$/))){
      const nm = m2[1];
      if(/^(DE )?(INGRESOS|GASTOS|EGRESOS|APORTES)$/.test(nm)) continue;
      if(/GASTOS OPERACIONALES/.test(nm)) continue;
      if(/GENERAL/.test(nm) && /GASTOS/.test(nm)) continue;
      if(/^GRAN/.test(lb)) continue;
      nombre = nm.replace(/^(DE|DEL)\s+/,'');
    }
    else if(lb.startsWith('SEGURO TODO RIESGO')) nombre = 'SEGURO TODO RIESGO';
    else if(lb === 'DIVERSOS') nombre = 'DIVERSOS';
    else if(lb.startsWith('FONDO DE IMPREVISTOS')) nombre = 'FONDO DE IMPREVISTOS';
    if(!nombre) continue;
    const e = num((rows[i]||[])[ejecCol]) || 0;
    if(e > 0) composicion.push({ categoria: tituloCorto(nombre), valor: e });
  }
  const gastosEjec = val(iGastosTotal, ejecCol);
  const sumComp = composicion.reduce((s, c) => s + c.valor, 0);
  if(!composicion.length || Math.abs(sumComp - gastosEjec) / (gastosEjec || 1) > 0.05){
    composicion = composicionDesdeDetalle(desviaciones, gastosEjec);
  }

  return {
    mes,
    presupuesto: {
      ingresos: { ppto: val(iIngresos, pptoCol), ejecutado: val(iIngresos, ejecCol) },
      gastos:   { ppto: val(iGastosTotal, pptoCol), ejecutado: gastosEjec }
    },
    excedenteAcumulado: iExcedente > 0 ? val(iExcedente, ejecCol)
      : (val(iIngresos, ejecCol) - gastosEjec),
    excedenteMes: iExcedente > 0 && mes ? val(iExcedente, monthCols[mes]) : null,
    mensual, desviaciones, composicion, ingresosDetalle, servicios
  };
}

function composicionDesdeDetalle(desviaciones, total){
  const conGasto = desviaciones.filter(d => d.ejecutado > 0).sort((a, b) => b.ejecutado - a.ejecutado);
  const top = conGasto.slice(0, 6).map(d => ({ categoria: d.rubro, valor: d.ejecutado }));
  const resto = (total || conGasto.reduce((s, d) => s + d.ejecutado, 0)) - top.reduce((s, c) => s + c.valor, 0);
  if(resto > 0) top.push({ categoria: 'Otros Rubros', valor: resto });
  return top;
}

// ── FORMATO B: ejecución anual + estado de resultados (Oz) ───────────────────
function parseFormatoB(wb){
  const erName = wb.SheetNames.find(n => /RESULTADOS/.test(norm(n)));
  const ejName = wb.SheetNames.find(n => /EJECUCION|EJEC/.test(norm(n)) && !/RESULTADOS/.test(norm(n)));
  if(!erName || !ejName) return null;
  const er = rowsOf(wb, erName);
  const ej = rowsOf(wb, ejName);

  // --- Estado de Resultados: DETALLE / ACUMULADO / MES ---
  const hER = findRow(er, r => cellMatch(r, /^ACUMULADO$/) >= 0);
  if(hER < 0) return null;
  const acumCol = cellMatch(er[hER], /^ACUMULADO$/);
  const mesCol = cellMatch(er[hER], /^MES$/);
  const lblCol = cellMatch(er[hER], /^DETALLE$/) >= 0 ? cellMatch(er[hER], /^DETALLE$/) : 1;

  // Mes actual desde los títulos ("Mayo 31 de 2026")
  let mes = null;
  for(let i = 0; i < Math.min(10, er.length); i++){
    for(const c of (er[i] || [])){ const m = mesEnTexto(c); if(m){ mes = m; break; } }
    if(mes) break;
  }
  const nMeses = mes ? MESES.indexOf(mes) + 1 : null;

  const lbl = (rows, i) => norm((rows[i] || [])[lblCol]);
  const iIng = findRowLast(er, r => /^TOTAL (DE )?INGRESOS$/.test(norm(r[lblCol])));
  const iEgr = findRow(er, r => /^TOTAL (DE )?(EGRESOS|GASTOS)$/.test(norm(r[lblCol])));
  let iExc = findRowLast(er, r => /^EXCEDENTES? DEL EJERCICIO/.test(norm(r[lblCol])));
  if(iExc < 0) iExc = findRowLast(er, r => /^RESULTADO DEL EJERCICIO/.test(norm(r[lblCol])));
  const iEgrHdr = findRow(er, r => /^(EGRESOS|GASTOS)\s*$/.test(norm(r[lblCol])));
  const iNoOp = findRow(er, r => /^(EGRESOS|GASTOS) NO OPERACIONALES/.test(norm(r[lblCol])));
  if(iIng < 0 || iEgr < 0) return null;

  const vER = (i, j) => num((er[i] || [])[j]) || 0;
  const noOpAcum = iNoOp >= 0 ? vER(iNoOp, acumCol) : 0;
  const noOpMes = iNoOp >= 0 && mesCol >= 0 ? vER(iNoOp, mesCol) : 0;
  const ingresosEjec = vER(iIng, acumCol);
  const gastosEjec = vER(iEgr, acumCol) + noOpAcum;

  // Composición: categorías entre EGRESOS y TOTAL EGRESOS + no operacionales
  const composicion = [];
  if(iEgrHdr >= 0){
    for(let i = iEgrHdr + 1; i < iEgr; i++){
      const l = lbl(er, i);
      if(!l || /^TOTAL/.test(l)) continue;
      const v = vER(i, acumCol);
      if(v > 0) composicion.push({ categoria: tituloCorto(er[i][lblCol]), valor: v });
    }
  }
  if(noOpAcum > 0) composicion.push({ categoria: 'Gastos No Operacionales', valor: noOpAcum });

  // Detalle de ingresos: filas hoja (leaf) antes de TOTAL INGRESOS
  const ingresosDetalle = [];
  for(let i = hER + 1; i < iIng; i++){
    const l = lbl(er, i);
    if(!l || /^TOTAL/.test(l)) continue;
    if(/^(INGRESOS|OTROS INGRESOS|APORTES)$/.test(l)) continue; // agregadores
    const v = vER(i, acumCol);
    if(v === 0) continue;
    ingresosDetalle.push({ rubro: tituloCorto(er[i][lblCol]), ppto: null, ejecutado: v, diferencia: null });
  }

  // --- Ejecución anual: CUENTA / PRESUPUESTO / EJECUCION → prorratear a la fecha ---
  const hEJ = findRow(ej, r => cellMatch(r, /PRESUPUESTO/) >= 0 && cellMatch(r, /EJECUCION/) >= 0);
  let desviaciones = [], gastosPpto = 0;
  if(hEJ >= 0 && nMeses){
    const cCta = cellMatch(ej[hEJ], /CUENTA|DETALLE|RUBRO/) >= 0 ? cellMatch(ej[hEJ], /CUENTA|DETALLE|RUBRO/) : 1;
    const cPpto = cellMatch(ej[hEJ], /PRESUPUESTO/);
    const cEjec = cellMatch(ej[hEJ], /^EJECUCION/);
    let totalAnual = 0;
    for(let i = hEJ + 1; i < ej.length; i++){
      const l = norm((ej[i] || [])[cCta]);
      const pAnual = num((ej[i] || [])[cPpto]), e = num((ej[i] || [])[cEjec]);
      if(!l){
        // fila de totales sin etiqueta: se ignora (el total se calcula sumando)
        continue;
      }
      if(/^TOTAL/.test(l) || /^AL HABER/.test(l)) break;
      if(pAnual === null && e === null) continue;
      const pProrr = Math.round((pAnual || 0) * nMeses / 12);
      totalAnual += (pAnual || 0);
      const ejec = e || 0;
      if(pProrr === 0 && ejec === 0) continue;
      desviaciones.push({ rubro: tituloCorto(ej[i][cCta]), ppto: pProrr, pptoAnual: pAnual || 0, ejecutado: ejec, diferencia: ejec - pProrr });
    }
    gastosPpto = Math.round(totalAnual * nMeses / 12);
  }

  return {
    mes,
    presupuesto: {
      // presupuesto de PH balanceado: ingresos presupuestados = gastos presupuestados
      ingresos: { ppto: gastosPpto, ejecutado: ingresosEjec },
      gastos:   { ppto: gastosPpto, ejecutado: gastosEjec }
    },
    excedenteAcumulado: iExc >= 0 ? vER(iExc, acumCol) : (ingresosEjec - gastosEjec),
    excedenteMes: iExc >= 0 && mesCol >= 0 ? vER(iExc, mesCol) : null,
    mensual: parseMensualB(wb) || (mes && mesCol >= 0 ? {
      meses: [mes],
      ingresos: [vER(iIng, mesCol)],
      gastos: [vER(iEgr, mesCol) + noOpMes]
    } : { meses: [], ingresos: [], gastos: [] }),
    desviaciones,
    composicion: composicion.length ? composicion : composicionDesdeDetalle(desviaciones, gastosEjec),
    ingresosDetalle,
    servicios: parseServiciosB(wb)
  };
}

// Hoja "PPTO" (formato Oz): matriz mensual — total de gastos y superávit por mes
function parseMensualB(wb){
  const name = wb.SheetNames.find(n => (/^P[PT]TO( |$)|PRESUPUESTO/.test(norm(n))) && !/EJEC/.test(norm(n)));
  if(!name) return null;
  const rows = rowsOf(wb, name);
  const mIdx = findRow(rows, r => r.filter(c => MESES.some(m => norm(c).startsWith(m))).length >= 3);
  if(mIdx < 0) return null;
  const monthCols = {};
  rows[mIdx].forEach((c, j) => {
    const n = norm(c);
    for(const m of MESES){ if(n.startsWith(m) && monthCols[m] === undefined) monthCols[m] = j; }
  });
  const meses = MESES.filter(m => monthCols[m] !== undefined);
  const lab = i => norm((rows[i] || [])[0]);
  const iGas = findRowLast(rows, r => /^(TOTAL (DE )?(GASTOS|EGRESOS|GTOS)|GASTOS OPERATIVOS$)/.test(norm(r[0])));
  if(iGas < 0) return null;
  const iSup = findRowLast(rows, r => /^(SUPERAVIT|EXCEDENTE|RESULTADO)/.test(norm(r[0])));
  const iIng = findRow(rows, r => /^TOTAL (DE )?(INGRESOS|APORTES)$/.test(norm(r[0])));
  const val = (i, j) => num((rows[i] || [])[j]) || 0;

  const gastos = meses.map(m => val(iGas, monthCols[m]));
  let ingresos = null;
  if(iSup > 0) ingresos = meses.map((m, x) => gastos[x] + val(iSup, monthCols[m]));
  else if(iIng > 0) ingresos = meses.map(m => val(iIng, monthCols[m]));
  if(!ingresos) return null;
  if(!gastos.some(v => v) || !ingresos.some(v => v)) return null;
  return { meses, ingresos, gastos, excedenteMes: iSup > 0 ? meses.map(m => val(iSup, monthCols[m])) : null };
}

// Hoja "SVC PUBLICOS" (formato Oz): filas por mes con acueducto/energía/gas
function parseServiciosB(wb){
  const name = wb.SheetNames.find(n => /SVC|SERV/.test(norm(n)) && /P/.test(norm(n)));
  if(!name) return [];
  const rows = rowsOf(wb, name);
  const hIdx = findRow(rows, r => cellMatch(r, /^MES$/) >= 0 && (cellMatch(r, /^M3$/) >= 0 || cellMatch(r, /^KW/) >= 0));
  if(hIdx < 0) return [];
  const h = rows[hIdx];
  const mesCol = cellMatch(h, /^MES$/);
  const m3Col = cellMatch(h, /^M3$/), kwCol = cellMatch(h, /^KW/), gasCol = cellMatch(h, /^GAS/);
  const valorAcue = m3Col >= 0 ? cellMatch(h.map((c, j) => j > m3Col && j < (kwCol > 0 ? kwCol : h.length) ? c : null), /VALOR/) : -1;
  const valorEner = kwCol >= 0 ? cellMatch(h.map((c, j) => j > kwCol ? c : null), /VALOR/) : -1;

  const meses = [], acue = [], ener = [], gas = [];
  for(let i = hIdx + 1; i < rows.length; i++){
    const r = rows[i] || [];
    const d = r[mesCol];
    let mesNombre = null;
    if(d instanceof Date) mesNombre = MESES[d.getMonth()];
    else if(typeof d === 'number' && d > 40000) mesNombre = MESES[new Date(Math.round((d - 25569) * 86400 * 1000)).getUTCMonth()];
    else mesNombre = mesEnTexto(d);
    if(!mesNombre) continue;
    meses.push(mesNombre.charAt(0) + mesNombre.slice(1).toLowerCase());
    acue.push(valorAcue >= 0 ? (num(r[valorAcue]) || 0) : null);
    ener.push(valorEner >= 0 ? (num(r[valorEner]) || 0) : null);
    gas.push(gasCol >= 0 ? (num(r[gasCol]) || 0) : null);
  }
  if(!meses.length) return [];
  const out = [];
  const mesesUp = meses.map(m => m.toUpperCase());
  if(valorAcue >= 0 && acue.some(v => v)) out.push({ nombre: 'Acueducto Y Saneamiento', meses: mesesUp, valores: acue });
  if(valorEner >= 0 && ener.some(v => v)) out.push({ nombre: 'Energía Eléctrica', meses: mesesUp, valores: ener });
  if(gasCol >= 0 && gas.some(v => v)) out.push({ nombre: 'Gas', meses: mesesUp, valores: gas });
  return out;
}

// ── Balance / Estado de Situación Financiera: efectivo disponible ────────────
function parseEfectivo(wb, mes){
  if(!mes) return null;
  const name = wb.SheetNames.find(n => /SITUAC|BALANCE|^BCE$|SIT\.? ?FRA/.test(norm(n)));
  if(!name) return null;
  const rows = rowsOf(wb, name);

  // Columna del mes actual: fila con varios meses, o celda corta que contenga el mes
  let mesCol = -1;
  const iMeses = findRow(rows, r => r.filter(c => MESES.some(m => norm(c).startsWith(m))).length >= 6);
  if(iMeses >= 0) mesCol = cellMatch(rows[iMeses], new RegExp('^' + mes + '\\b'));
  if(mesCol < 0){
    const re = new RegExp('(^|\\b)' + mes + '\\b');
    for(let i = 0; i < Math.min(15, rows.length) && mesCol < 0; i++){
      const r = rows[i] || [];
      for(let j = 1; j < r.length; j++){
        const n = norm(r[j]);
        if(n && n.length <= 25 && re.test(n)){ mesCol = j; break; }
      }
    }
  }
  if(mesCol < 0) return null;

  const iEf = findRow(rows, r => cellMatch(r, /^(EFECTIVO Y|CAJA Y BANCOS)/) >= 0);
  if(iEf < 0) return null;
  // 1) valor en la propia fila; 2) columna de totales de grupo (mesCol+1); 3) fila subtotal debajo
  let v = num((rows[iEf] || [])[mesCol]);
  if(v !== null && v !== 0) return v;
  v = num((rows[iEf] || [])[mesCol + 1]);
  if(v !== null && v !== 0) return v;
  for(let i = iEf + 1; i <= iEf + 8 && i < rows.length; i++){
    const r = rows[i] || [];
    const vi = num(r[mesCol]);
    if(vi === null) continue;
    const lb = norm(r.find(c => typeof c === 'string' && norm(c).length > 2) || '');
    const tieneEtiqueta = r.slice(0, mesCol).some(c => typeof c === 'string' && norm(c).length > 2);
    if(!tieneEtiqueta || /^TOTAL (EFECTIVO|CAJA)/.test(lb)) return vi;
  }
  return null;
}

// ── Cartera / cuentas por cobrar ─────────────────────────────────────────────
function parseCartera(wb){
  const name = wb.SheetNames.find(n => /COBRAR|CARTERA|C ?X ?C|CTAS ?X ?C\b/.test(norm(n)) && !/PAGAR|C ?X ?P\b|GRAFICO|FACTURACION/.test(norm(n)));
  if(!name) return { carteraTotal: null, morosos: [] };
  const rows = rowsOf(wb, name);

  const hIdx = findRow(rows, r => cellMatch(r, /VALOR (DE CUOTAS )?VENCID|^SALDO( ACTUAL)?$|^VALOR *\$?$/) >= 0);
  if(hIdx < 0) return { carteraTotal: null, morosos: [] };
  const header = rows[hIdx];
  const valCol = cellMatch(header, /VALOR (DE CUOTAS )?VENCID|^SALDO( ACTUAL)?$|^VALOR *\$?$/);
  const cuotasCol = cellMatch(header, /#.*CUOTAS/);

  const porApto = {};
  for(let i = hIdx + 1; i < rows.length; i++){
    const r = rows[i] || [];
    const valor = num(r[valCol]);
    if(valor === null) continue;
    // id del inmueble: cualquier celda antes de valCol que sea entero o código corto
    let idCol = -1, id = null, nombre = '';
    for(let j = 0; j < valCol; j++){
      const raw = r[j];
      if(raw === null || raw === '') continue;
      const s = String(raw).trim();
      const n = norm(s);
      if(/^\d{1,6}$/.test(s)){ idCol = j; id = s; break; }
      if(typeof raw === 'string' && s.length <= 8 && /^[A-ZÑ0-9\/\.\-]{2,8}$/.test(n) && !/^(TOTAL|NOTA)/.test(n) && !MESES.includes(n)){ idCol = j; id = s; }
    }
    if(idCol < 0 || valor <= 0) continue;
    // nombre: la celda de texto más larga (≠ id) antes de valCol
    for(let j = 0; j < valCol; j++){
      if(j === idCol) continue;
      const c = r[j];
      if(typeof c === 'string' && c.trim().length > nombre.length) nombre = c.replace(/\s+/g,' ').trim();
    }
    if(!porApto[id]) porApto[id] = { apto: id, nombre, valor: 0, cuotas: null };
    porApto[id].valor += valor;
    if(!porApto[id].nombre && nombre) porApto[id].nombre = nombre;
    const c = cuotasCol >= 0 ? num(r[cuotasCol]) : null;
    if(c !== null && porApto[id].cuotas === null) porApto[id].cuotas = Math.round(c * 10) / 10;
  }
  const morosos = Object.values(porApto).sort((a, b) => b.valor - a.valor);

  // Total: fila "TOTAL CARTERA ..." o "Totales"; si no existe, suma de morosos
  let carteraTotal = null;
  const t = findLabelCell(rows, /^TOTAL CARTERA|^TOTALES$/);
  if(t) carteraTotal = num(rows[t.i][valCol]);
  if(carteraTotal === null) carteraTotal = morosos.reduce((s, m) => s + m.valor, 0);
  return { carteraTotal, morosos };
}

// ── Parser de archivo completo ───────────────────────────────────────────────
function parseArchivo(buffer, filename){
  const wb = XLSX.read(buffer, { type: 'buffer' });

  // Formato A: buscar una hoja mensualizada de ejecución
  let ejec = null;
  for(const n of wb.SheetNames){
    if(!/EJEC/.test(norm(n))) continue;
    ejec = parseFormatoA(rowsOf(wb, n));
    if(ejec) break;
  }
  // Formato B (Oz)
  if(!ejec) ejec = parseFormatoB(wb);
  if(!ejec) throw new Error(filename + ': formato de ejecución presupuestal no reconocido');

  const efectivo = parseEfectivo(wb, ejec.mes);
  const cxc = parseCartera(wb);

  const base = path.basename(filename).replace(/\.(xlsx?|xlsm)$/i, '');
  const partes = base.split(/[-_]/).filter(p => p && norm(p) !== 'EEFF' && !/^(19|20)\d{2}$/.test(p));
  const anioM = base.match(/(19|20)\d{2}/);
  const nMeses = ejec.mes ? MESES.indexOf(ejec.mes) + 1 : null;

  return {
    unidad: tituloCorto(partes.join(' ')) || base,
    archivo: path.basename(filename),
    anio: anioM ? parseInt(anioM[0]) : null,
    mes: ejec.mes,
    kpis: {
      efectivo,
      cartera: cxc.carteraTotal,
      excedenteAcumulado: ejec.excedenteAcumulado,
      excedenteMes: ejec.excedenteMes !== undefined && ejec.excedenteMes !== null ? ejec.excedenteMes
        : (ejec.mensual && ejec.mensual.excedenteMes ? ejec.mensual.excedenteMes[ejec.mensual.excedenteMes.length - 1] : null),
      gastoPromedioMensual: nMeses ? ejec.presupuesto.gastos.ejecutado / nMeses : null
    },
    presupuesto: ejec.presupuesto,
    mensual: ejec.mensual,
    morosos: cxc.morosos,
    desviaciones: ejec.desviaciones,
    composicion: ejec.composicion,
    ingresosDetalle: ejec.ingresosDetalle || [],
    servicios: ejec.servicios || []
  };
}

// ── Microsoft Graph ──────────────────────────────────────────────────────────
async function getToken(fetch){
  const url = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const r = await fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
  if(!r.ok) throw new Error(`Token failed: ${r.status} — ${await r.text()}`);
  return (await r.json()).access_token;
}

function shareId(url){
  const b64 = Buffer.from(url).toString('base64').replace(/=+$/,'').replace(/\//g,'_').replace(/\+/g,'-');
  return 'u!' + b64;
}

async function listarCarpetaEEFF(fetch, token){
  const H = { Authorization: `Bearer ${token}` };
  try{
    const r = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId(EEFF_SHARE_URL)}/driveItem?$expand=children`, { headers: H });
    if(r.ok){
      const item = await r.json();
      console.log(`      Carpeta resuelta vía link compartido: "${item.name}" (${(item.children||[]).length} elementos)`);
      return { children: item.children || [] };
    }
    console.log(`      Link compartido no resoluble (${r.status}), probando rutas directas...`);
  }catch(e){ console.log('      Link compartido falló:', e.message); }

  const rs = await fetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`, { headers: H });
  if(!rs.ok) throw new Error(`Site lookup failed: ${rs.status}`);
  const siteId = (await rs.json()).id;
  for(const p of FOLDER_PATH_CANDIDATES){
    const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:${encodeURI(p)}:/children`, { headers: H });
    if(r.ok){
      const json = await r.json();
      console.log(`      Carpeta encontrada en ruta directa: ${p} (${json.value.length} elementos)`);
      return { children: json.value };
    }
  }
  throw new Error('No se pudo encontrar la carpeta EEFF ni por link compartido ni por rutas directas. ' +
    'Verifica permisos de la app (Sites.Read.All / Files.Read.All) o agrega la ruta correcta en FOLDER_PATH_CANDIDATES.');
}

async function descargarItem(fetch, token, item){
  const url = item['@microsoft.graph.downloadUrl'] ||
    `https://graph.microsoft.com/v1.0/drives/${item.parentReference.driveId}/items/${item.id}/content`;
  const headers = item['@microsoft.graph.downloadUrl'] ? {} : { Authorization: `Bearer ${token}` };
  const r = await fetch(url, { headers });
  if(!r.ok) throw new Error(`Download ${item.name} failed: ${r.status}`);
  return Buffer.from(await r.arrayBuffer());
}

// ── Main ─────────────────────────────────────────────────────────────────────
(async () => {
  try{
    const localIdx = process.argv.indexOf('--local');
    let unidades = [];

    if(localIdx > 0){
      for(const f of process.argv.slice(localIdx + 1)){
        try{
          const u = parseArchivo(fs.readFileSync(f), f);
          console.log(`✓ ${path.basename(f)} → ${u.unidad} (corte: ${u.mes})`);
          unidades.push(u);
        }catch(e){
          console.error(`✗ ${path.basename(f)}: ${e.message}`);
        }
      }
    }else{
      const fetch = require('node-fetch');
      if(!TENANT || !CLIENT_ID || !CLIENT_SECRET){
        console.error('ERROR: Faltan secrets de Azure.'); process.exit(1);
      }
      console.log('[1/4] Obteniendo token de Microsoft...');
      const token = await getToken(fetch);
      console.log('      OK');

      console.log('[2/4] Listando carpeta EEFF...');
      const { children } = await listarCarpetaEEFF(fetch, token);
      // Orden ascendente por fecha de modificación: en caso de duplicados gana el más reciente
      const excels = children
        .filter(c => c.file && /\.(xlsx?|xlsm)$/i.test(c.name) && !c.name.startsWith('~$'))
        .sort((a, b) => new Date(a.lastModifiedDateTime || 0) - new Date(b.lastModifiedDateTime || 0));
      if(!excels.length) throw new Error('La carpeta EEFF no contiene archivos Excel.');
      console.log('      Archivos:', excels.map(e => e.name).join(', '));

      console.log('[3/4] Descargando y parseando...');
      for(const item of excels){
        try{
          const buf = await descargarItem(fetch, token, item);
          const u = parseArchivo(buf, item.name);
          console.log(`      ✓ ${item.name} → ${u.unidad} (corte: ${u.mes})`);
          unidades.push(u);
        }catch(e){
          console.error(`      ✗ ${item.name}: ${e.message}`);
        }
      }
      if(!unidades.length) throw new Error('Ningún archivo se pudo parsear.');
    }

    // Deduplicar por unidad+año (ej. mismo EEFF en .xls y .xlsx): gana el último parseado
    const dedup = {};
    for(const u of unidades) dedup[u.unidad + '|' + u.anio] = u;
    unidades = Object.values(dedup).sort((a, b) => a.unidad.localeCompare(b.unidad));

    console.log('[4/4] Escribiendo finanzas.json...');
    const out = { actualizado: new Date().toISOString(), total: unidades.length, unidades };
    fs.writeFileSync('finanzas.json', JSON.stringify(out, null, 2));
    console.log(`      ${unidades.length} unidad(es): ${unidades.map(u => u.unidad).join(', ')}`);
    console.log('LISTO.');
  }catch(e){
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
