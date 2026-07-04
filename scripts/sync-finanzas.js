// scripts/sync-finanzas.js
// Descarga los archivos EEFF-*.xls de la carpeta de SharePoint (vía Microsoft Graph),
// extrae KPIs financieros por unidad y los guarda en finanzas.json
//
// Uso normal (GitHub Actions):  node scripts/sync-finanzas.js
// Uso local (probar parser):    node scripts/sync-finanzas.js --local ruta/al/EEFF-UNIDAD-2026.xls

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const TENANT = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// === CONFIGURACIÓN DE SHAREPOINT ===
const SHAREPOINT_HOST = 'escalarasesorias.sharepoint.com';
const SITE_PATH = '/sites/Escalar';
// Link compartido de la carpeta EEFF (se resuelve vía Graph /shares)
const EEFF_SHARE_URL = 'https://escalarasesorias.sharepoint.com/:f:/s/Escalar/IgAFNxZ-bfPtRrpMDL4yJwEIATOW2G4xvwKxEa7phg1yGQc?e=89GeKh';
// Rutas alternativas por si el link compartido no se puede resolver con permisos de aplicación
const FOLDER_PATH_CANDIDATES = ['/Escalar/EEFF', '/EEFF', '/Escalar/Estados Financieros'];
// ===================================

const MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE'];

// ── Utilidades ────────────────────────────────────────────────────────────────
function norm(s){ return String(s == null ? '' : s).replace(/\s+/g,' ').trim().toUpperCase(); }
function num(v){ const n = typeof v === 'number' ? v : parseFloat(String(v||'').replace(/[^0-9.\-]/g,'')); return isFinite(n) ? n : null; }
function sheetToRows(wb, namePart){
  const sheetName = wb.SheetNames.find(n => norm(n).includes(namePart));
  if(!sheetName) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null });
}
function findRow(rows, pred){
  for(let i = 0; i < rows.length; i++){ if(pred(rows[i] || [], i)) return i; }
  return -1;
}
function cellMatch(row, re){
  for(let j = 0; j < row.length; j++){ if(re.test(norm(row[j]))) return j; }
  return -1;
}

// ── Parser: EJEC PPTAL ───────────────────────────────────────────────────────
function parseEjecPptal(rows){
  const hIdx = findRow(rows, r => r.some(c => norm(c) === 'DETALLE'));
  if(hIdx < 0) throw new Error('EJEC PPTAL: no se encontró fila de encabezado DETALLE');
  const header = rows[hIdx];

  const monthCols = {};
  header.forEach((c, j) => { const n = norm(c); if(MESES.includes(n)) monthCols[n] = j; });

  const pptoCol = cellMatch(header, /^PTTO/);
  const ejecCol = cellMatch(header, /^EJECUCI/);
  const difCol  = cellMatch(header, /^DIFERENCIA/);
  if(pptoCol < 0 || ejecCol < 0) throw new Error('EJEC PPTAL: no se encontraron columnas PTTO/EJECUCIÓN');

  const mMatch = norm(header[ejecCol]).match(/EJECUCI[OÓ]N A (\w+)/);
  const mesActual = mMatch && MESES.includes(mMatch[1]) ? mMatch[1] : null;

  const label = i => norm((rows[i] || [])[0]);
  const iIngresos = findRow(rows, r => norm(r[0]) === 'TOTAL INGRESOS');
  const iGranTotal = findRow(rows, r => norm(r[0]).startsWith('GRAN TOTAL DE GASTOS'));
  const iTotalGral = findRow(rows, r => norm(r[0]).startsWith('TOTAL GENERAL GASTOS'));
  const iGastos = findRow(rows, r => norm(r[0]) === 'GASTOS');
  const iExcedente = findRow(rows, r => norm(r[0]).startsWith('EXCEDENTE O'));
  if(iIngresos < 0 || iGranTotal < 0) throw new Error('EJEC PPTAL: no se encontraron filas de totales');

  const val = (i, j) => num((rows[i] || [])[j]) || 0;

  // Serie mensual hasta el mes actual
  const mesesHasta = mesActual ? MESES.slice(0, MESES.indexOf(mesActual) + 1) : [];
  const mensual = {
    meses: mesesHasta,
    ingresos: mesesHasta.map(m => val(iIngresos, monthCols[m])),
    gastos: mesesHasta.map(m => val(iGranTotal, monthCols[m]))
  };

  // Detalle de rubros de gasto (para tornado): entre GASTOS y fin de la sección
  const desviaciones = [];
  const fin = Math.max(iExcedente > 0 ? iExcedente : rows.length, iGranTotal + 1);
  for(let i = (iGastos > 0 ? iGastos + 1 : hIdx + 1); i < fin; i++){
    const lb = label(i);
    if(!lb) continue;
    if(/^(SUBTOTAL|TOTAL|GRAN TOTAL|CIFRAS|EXCEDENTE)/.test(lb)) continue;
    if(lb === 'DIVERSOS') continue; // agrupador: sus detalles vienen en las filas siguientes
    const p = num((rows[i]||[])[pptoCol]), e = num((rows[i]||[])[ejecCol]);
    if(p === null && e === null) continue;            // encabezado de sección
    const ppto = p || 0, ejec = e || 0;
    if(ppto === 0 && ejec === 0) continue;
    desviaciones.push({ rubro: tituloCorto(rows[i][0]), ppto, ejecutado: ejec, diferencia: ejec - ppto });
  }

  // Composición del gasto: subtotales + rubros de primer nivel
  const composicion = [];
  for(let i = hIdx + 1; i < rows.length; i++){
    const lb = label(i);
    let nombre = null;
    if(lb.startsWith('SUBTOTAL')) nombre = lb.replace(/^SUBTOTAL(ES)?\s*/,'');
    else if(lb.startsWith('SEGURO TODO RIESGO')) nombre = 'SEGURO TODO RIESGO';
    else if(lb === 'DIVERSOS') nombre = 'DIVERSOS';
    else if(lb.startsWith('FONDO DE IMPREVISTOS')) nombre = 'FONDO DE IMPREVISTOS';
    if(!nombre) continue;
    const e = num((rows[i]||[])[ejecCol]) || 0;
    if(e > 0) composicion.push({ categoria: tituloCorto(nombre), valor: e });
  }

  return {
    mes: mesActual,
    presupuesto: {
      ingresos: { ppto: val(iIngresos, pptoCol), ejecutado: val(iIngresos, ejecCol) },
      gastos:   { ppto: val(iGranTotal, pptoCol), ejecutado: val(iGranTotal, ejecCol) }
    },
    excedenteAcumulado: iExcedente > 0 ? val(iExcedente, ejecCol) : (val(iIngresos, ejecCol) - val(iGranTotal, ejecCol)),
    mensual, desviaciones, composicion
  };
}
function tituloCorto(s){
  return String(s||'').replace(/\s+/g,' ').trim()
    .toLowerCase().replace(/(^|\s|\/|-)\S/g, t => t.toUpperCase());
}

// ── Parser: EST DE SITUACION FRA ─────────────────────────────────────────────
function parseBalance(rows, mesActual){
  const iMeses = findRow(rows, r => cellMatch(r, /^ENERO$/) >= 0 && cellMatch(r, /^DICIEMBRE$/) >= 0);
  const iEfectivo = findRow(rows, r => cellMatch(r, /^EFECTIVO Y EQUIVALENTE/) >= 0);
  if(iMeses < 0 || iEfectivo < 0 || !mesActual) return { efectivo: null };
  const col = cellMatch(rows[iMeses], new RegExp('^' + mesActual + '$'));
  if(col < 0) return { efectivo: null };
  return { efectivo: num(rows[iEfectivo][col]) };
}

// ── Parser: ANEXOS C X C ─────────────────────────────────────────────────────
function parseCartera(rows){
  const hIdx = findRow(rows, r => cellMatch(r, /VALOR VENCIDO/) >= 0);
  if(hIdx < 0) return { carteraTotal: null, morosos: [] };
  const valCol = cellMatch(rows[hIdx], /VALOR VENCIDO/);
  const cuotasCol = cellMatch(rows[hIdx], /CUOTAS VENCIDAS/);
  const nombreCol = cellMatch(rows[hIdx], /CUOTAS DE ADMINISTRACION/);

  const porApto = {};
  for(let i = hIdx + 1; i < rows.length; i++){
    const r = rows[i] || [];
    const apto = num(r[0]);
    if(apto === null || !Number.isInteger(apto)) continue;
    const valor = num(r[valCol]);
    if(valor === null || valor === 0) continue;
    const key = String(apto);
    if(!porApto[key]) porApto[key] = { apto: key, nombre: String(r[nombreCol >= 0 ? nombreCol : 1] || '').replace(/\s+/g,' ').trim(), valor: 0, cuotas: null };
    porApto[key].valor += valor;
    const c = cuotasCol >= 0 ? num(r[cuotasCol]) : null;
    if(c !== null && porApto[key].cuotas === null) porApto[key].cuotas = Math.round(c * 10) / 10;
  }
  const morosos = Object.values(porApto).sort((a, b) => b.valor - a.valor);

  const iTotal = findRow(rows, r => cellMatch(r, /^TOTAL CARTERA/) >= 0);
  const carteraTotal = iTotal >= 0 ? num(rows[iTotal][cellMatch(rows[iTotal], /^TOTAL CARTERA/) >= 0 ? valCol : valCol]) : null;
  return { carteraTotal, morosos };
}

// ── Parser: archivo completo ─────────────────────────────────────────────────
function parseArchivo(buffer, filename){
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ejecRows = sheetToRows(wb, 'EJEC');
  const balRows = sheetToRows(wb, 'SITUACION');
  const cxcRows = sheetToRows(wb, 'C X C') || sheetToRows(wb, 'CXC');
  if(!ejecRows) throw new Error(filename + ': no se encontró hoja de ejecución presupuestal');

  const ejec = parseEjecPptal(ejecRows);
  const bal = balRows ? parseBalance(balRows, ejec.mes) : { efectivo: null };
  const cxc = cxcRows ? parseCartera(cxcRows) : { carteraTotal: null, morosos: [] };

  // Nombre de la unidad y año desde el nombre del archivo: EEFF-CHARLOTTE-2026.xls
  const base = path.basename(filename).replace(/\.(xlsx?|xlsm)$/i, '');
  const partes = base.split(/[-_]/).filter(p => p && norm(p) !== 'EEFF' && !/^(19|20)\d{2}$/.test(p));
  const anioM = base.match(/(19|20)\d{2}/);

  const gastoPromMes = ejec.mensual.gastos.length ? ejec.mensual.gastos.reduce((a,b)=>a+b,0) / ejec.mensual.gastos.length : null;

  return {
    unidad: tituloCorto(partes.join(' ')) || base,
    archivo: path.basename(filename),
    anio: anioM ? parseInt(anioM[0]) : null,
    mes: ejec.mes,
    kpis: {
      efectivo: bal.efectivo,
      cartera: cxc.carteraTotal,
      excedenteAcumulado: ejec.excedenteAcumulado,
      gastoPromedioMensual: gastoPromMes
    },
    presupuesto: ejec.presupuesto,
    mensual: ejec.mensual,
    morosos: cxc.morosos,
    desviaciones: ejec.desviaciones,
    composicion: ejec.composicion
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
  // 1) Intentar resolver el link compartido
  try{
    const r = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareId(EEFF_SHARE_URL)}/driveItem?$expand=children`, { headers: H });
    if(r.ok){
      const item = await r.json();
      console.log(`      Carpeta resuelta vía link compartido: "${item.name}" (${(item.children||[]).length} elementos)`);
      return { driveId: item.parentReference.driveId, children: item.children || [] };
    }
    console.log(`      Link compartido no resoluble (${r.status}), probando rutas directas...`);
  }catch(e){ console.log('      Link compartido falló:', e.message); }

  // 2) Rutas directas dentro del sitio
  const rs = await fetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`, { headers: H });
  if(!rs.ok) throw new Error(`Site lookup failed: ${rs.status}`);
  const siteId = (await rs.json()).id;
  for(const p of FOLDER_PATH_CANDIDATES){
    const r = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:${encodeURI(p)}:/children`, { headers: H });
    if(r.ok){
      const json = await r.json();
      console.log(`      Carpeta encontrada en ruta directa: ${p} (${json.value.length} elementos)`);
      return { siteId, children: json.value };
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
      // Modo local: parsear archivos dados por línea de comandos
      const files = process.argv.slice(localIdx + 1);
      for(const f of files){
        console.log('Parseando local:', f);
        unidades.push(parseArchivo(fs.readFileSync(f), f));
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
      const excels = children.filter(c => c.file && /\.(xlsx?|xlsm)$/i.test(c.name) && !c.name.startsWith('~$'));
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

    console.log('[4/4] Escribiendo finanzas.json...');
    unidades.sort((a, b) => a.unidad.localeCompare(b.unidad));
    const out = { actualizado: new Date().toISOString(), total: unidades.length, unidades };
    fs.writeFileSync('finanzas.json', JSON.stringify(out, null, 2));
    console.log(`      ${unidades.length} unidad(es) escritas en finanzas.json`);
    console.log('LISTO.');
  }catch(e){
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
