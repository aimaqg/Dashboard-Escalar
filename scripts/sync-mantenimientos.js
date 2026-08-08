// scripts/sync-mantenimientos.js
// Descarga Mantenimientos_Escalar.xlsx desde SharePoint y genera mantenimientos.json
// El parsing real lo hace parse-mantenimientos.py (maneja celdas fusionadas con openpyxl)
//
// Uso local: node scripts/sync-mantenimientos.js --local ruta/Mantenimientos_Escalar.xlsx
// Producción: node scripts/sync-mantenimientos.js  (descarga de SharePoint)

const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const TENANT        = process.env.AZURE_TENANT_ID;
const CLIENT_ID     = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

const SHAREPOINT_HOST    = 'escalarasesorias.sharepoint.com';
const SITE_PATH          = '/sites/Escalar';
const MTTO_FILENAME      = 'Mantenimientos_Escalar.xlsx';
const FOLDER_CANDIDATES  = ['/Escalar', '/Escalar/EEFF', '/'];
const PARSER_SCRIPT      = path.join(__dirname, 'parse-mantenimientos.py');

async function getToken(fetch){
  const r = await fetch(
    `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
    { method:'POST',
      headers:{'Content-Type':'application/x-www-form-urlencoded'},
      body: new URLSearchParams({client_id:CLIENT_ID,client_secret:CLIENT_SECRET,
        scope:'https://graph.microsoft.com/.default',grant_type:'client_credentials'}) }
  );
  if(!r.ok) throw new Error(`Token: ${r.status} — ${await r.text()}`);
  return (await r.json()).access_token;
}

async function descargarArchivo(fetch, token){
  const H = { Authorization:`Bearer ${token}` };
  const rs = await fetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:${SITE_PATH}`, {headers:H});
  if(!rs.ok) throw new Error(`Site: ${rs.status}`);
  const siteId = (await rs.json()).id;

  for(const folder of FOLDER_CANDIDATES){
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:${encodeURI(folder+'/'+MTTO_FILENAME)}:/content`;
    const r = await fetch(url, {headers:H});
    if(r.ok){
      console.log(`  Archivo: ${folder}/${MTTO_FILENAME}`);
      const tmpPath = path.join('/tmp', MTTO_FILENAME);
      fs.writeFileSync(tmpPath, Buffer.from(await r.arrayBuffer()));
      return tmpPath;
    }
  }
  throw new Error(`${MTTO_FILENAME} no encontrado. Carpetas probadas: ${FOLDER_CANDIDATES.join(', ')}`);
}

function parsearConPython(excelPath){
  // Modificar temporalmente el path en el script Python y correrlo
  const py = fs.readFileSync(PARSER_SCRIPT, 'utf8');
  const pyTemp = py.replace(
    /open\(['"].*?Mantenimientos_Escalar\.xlsx['"]\)/g,
    `open(${JSON.stringify(excelPath)})`
  ).replace(
    /wb = openpyxl\.load_workbook\(['"].*?['"]\)/,
    `wb = openpyxl.load_workbook(${JSON.stringify(excelPath)})`
  );
  const tmpPy = '/tmp/parse_mtto_run.py';
  // Cambiar la ruta de salida al directorio actual
  const pyFinal = pyTemp.replace(
    "'/home/claude/Dashboard-Escalar/mantenimientos.json'",
    "'mantenimientos.json'"
  );
  fs.writeFileSync(tmpPy, pyFinal);
  try{
    const out = execSync(`python3 ${tmpPy}`, {encoding:'utf8'});
    console.log(out.trim());
  } catch(e){
    throw new Error('parse-mantenimientos.py falló: ' + e.stderr);
  }
}

(async () => {
  try{
    const localIdx = process.argv.indexOf('--local');
    if(localIdx >= 0){
      const file = process.argv[localIdx+1];
      if(!file){ console.error('Uso: --local archivo.xlsx'); process.exit(1); }
      console.log('[LOCAL] Parseando:', file);
      parsearConPython(path.resolve(file));
    } else {
      const fetch = require('node-fetch');
      if(!TENANT||!CLIENT_ID||!CLIENT_SECRET){
        console.error('ERROR: Faltan secrets de Azure.'); process.exit(1);
      }
      console.log('[1/3] Token...');
      const token = await getToken(fetch);
      console.log('[2/3] Descargando desde SharePoint...');
      const tmpPath = await descargarArchivo(fetch, token);
      console.log('[3/3] Parseando con Python...');
      parsearConPython(tmpPath);
    }
    console.log('LISTO — mantenimientos.json generado');
  } catch(e){
    console.error('FALLÓ:', e.message);
    process.exit(1);
  }
})();
