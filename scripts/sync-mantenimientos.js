// scripts/sync-mantenimientos.js
// Descarga Mantenimientos_Escalar.xlsx desde SharePoint vía link compartido
// y genera mantenimientos.json usando parse-mantenimientos.py
//
// Uso local: node scripts/sync-mantenimientos.js --local ruta/archivo.xlsx
// Producción: node scripts/sync-mantenimientos.js

const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const TENANT        = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// Link compartido del archivo en SharePoint
const MTTO_SHARE_URL = 'https://escalarasesorias.sharepoint.com/:x:/s/Escalar/IQD6GBEzVlcKRqKOYfmMcPWBAUvxuC9YOu4mcl97IdZ4BzA?e=yWzsE6';
const PARSER_SCRIPT  = path.join(__dirname, 'parse-mantenimientos.py');

function shareId(url){
  const b64 = Buffer.from(url).toString('base64')
    .replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  return 'u!' + b64;
}

async function getToken(fetch){
  const r = await fetch(
    `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
    { method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: CLIENT_ID, client_secret: CLIENT_SECRET,
        scope: 'https://graph.microsoft.com/.default',
        grant_type: 'client_credentials'
      })
    }
  );
  if(!r.ok) throw new Error(`Token: ${r.status} — ${await r.text()}`);
  return (await r.json()).access_token;
}

async function descargarArchivo(fetch, token){
  const sid = shareId(MTTO_SHARE_URL);
  const url = `https://graph.microsoft.com/v1.0/shares/${sid}/driveItem`;
  console.log('  Resolviendo link compartido...');
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if(!r.ok) throw new Error(`No se pudo resolver el link de Mantenimientos (${r.status}). Verifica que el archivo siga en SharePoint.`);
  const item = await r.json();
  console.log(`  Archivo: ${item.name} (modificado: ${item.lastModifiedDateTime})`);

  // Descargar contenido
  const dlUrl = item['@microsoft.graph.downloadUrl'] ||
    `https://graph.microsoft.com/v1.0/drives/${item.parentReference.driveId}/items/${item.id}/content`;
  const headers = item['@microsoft.graph.downloadUrl'] ? {} : { Authorization: `Bearer ${token}` };
  const dl = await fetch(dlUrl, { headers });
  if(!dl.ok) throw new Error(`Descarga fallida: ${dl.status}`);

  const tmpPath = '/tmp/Mantenimientos_Escalar.xlsx';
  fs.writeFileSync(tmpPath, Buffer.from(await dl.arrayBuffer()));
  console.log(`  Guardado temporalmente en ${tmpPath}`);
  return tmpPath;
}

function parsearConPython(excelPath){
  const py = fs.readFileSync(PARSER_SCRIPT, 'utf8');
  // Sustituir la ruta del archivo y la ruta de salida
  const pyFinal = py
    .replace(/wb = openpyxl\.load_workbook\([^)]+\)/,
             `wb = openpyxl.load_workbook(${JSON.stringify(excelPath)})`)
    .replace(/open\([^)]*Mantenimientos_Escalar\.xlsx[^)]*\)/,
             `open(${JSON.stringify(excelPath)})`)
    .replace("'/home/claude/Dashboard-Escalar/mantenimientos.json'",
             "'mantenimientos.json'");
  const tmpPy = '/tmp/parse_mtto_run.py';
  fs.writeFileSync(tmpPy, pyFinal);
  const out = execSync(`python3 ${tmpPy}`, { encoding: 'utf8' });
  console.log(out.trim());
}

(async () => {
  try {
    const localIdx = process.argv.indexOf('--local');
    if(localIdx >= 0){
      const file = process.argv[localIdx + 1];
      if(!file){ console.error('Uso: --local archivo.xlsx'); process.exit(1); }
      console.log(`[LOCAL] Parseando: ${file}`);
      parsearConPython(path.resolve(file));
    } else {
      const fetch = require('node-fetch');
      if(!TENANT || !CLIENT_ID || !CLIENT_SECRET){
        console.error('ERROR: Faltan secrets de Azure.'); process.exit(1);
      }
      console.log('[1/3] Obteniendo token...');
      const token = await getToken(fetch);
      console.log('      OK');
      console.log('[2/3] Descargando Mantenimientos_Escalar.xlsx...');
      const tmpPath = await descargarArchivo(fetch, token);
      console.log('[3/3] Parseando...');
      parsearConPython(tmpPath);
    }
    console.log('LISTO — mantenimientos.json generado');
  } catch(e) {
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
