// scripts/sync.js
// Descarga Tareas.xlsx desde SharePoint (vía Microsoft Graph) y lo guarda como tareas.json

const fetch = require('node-fetch');
const XLSX = require('xlsx');
const fs = require('fs');

const TENANT = process.env.AZURE_TENANT_ID;
const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// === CONFIGURACIÓN DE SHAREPOINT ===
const SHAREPOINT_HOST = 'escalarasesorias.sharepoint.com';
const SITE_PATH = '/sites/Escalar';
// Ruta del archivo DENTRO de la biblioteca de documentos del sitio.
// El archivo está en la carpeta "Escalar" del sitio.
const FILE_PATH = '/Escalar/Registro_Tareas_Teams.xlsx';
// ===================================

if (!TENANT || !CLIENT_ID || !CLIENT_SECRET) {
  console.error('ERROR: Faltan secrets de Azure. Configura en GitHub Settings > Secrets.');
  process.exit(1);
}

async function getToken() {
  const url = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const r = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  if (!r.ok) {
    const err = await r.text();
    throw new Error(`Token failed: ${r.status} — ${err}`);
  }
  const json = await r.json();
  return json.access_token;
}

async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const err = await r.text();
    throw new Error(`Site lookup failed: ${r.status} — ${err}`);
  }
  const json = await r.json();
  return json.id;
}

async function downloadExcel(token, siteId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:${FILE_PATH}:/content`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const err = await r.text();
    throw new Error(`Download failed: ${r.status} — ${err}`);
  }
  const arrayBuffer = await r.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

(async () => {
  try {
    console.log('[1/4] Obteniendo token de Microsoft...');
    const token = await getToken();
    console.log('      OK');

    console.log('[2/4] Buscando site de SharePoint...');
    const siteId = await getSiteId(token);
    console.log(`      Site ID: ${siteId}`);

    console.log('[3/4] Descargando Excel...');
    const buffer = await downloadExcel(token, siteId);
    console.log(`      ${buffer.length} bytes descargados`);

    console.log('[4/4] Convirtiendo a JSON...');
    const wb = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    const out = {
      actualizado: new Date().toISOString(),
      total: data.length,
      datos: data
    };

    fs.writeFileSync('tareas.json', JSON.stringify(out, null, 2));
    console.log(`      ${data.length} tareas escritas en tareas.json`);
    console.log('LISTO.');
  } catch (e) {
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
